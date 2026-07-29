// This function serves your HTML file when you deploy as a Web App
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Phil-IRI Assessment Tool')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Automatically runs when the Google Sheet is opened natively
function onOpen(e) {
  setupSpreadsheet();
  
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Phil-IRI Tool')
      .addItem('Setup / Refresh Tabs', 'setupSpreadsheet')
      .addItem('Forced Sync (Master Log -> Grade Tabs)', 'syncMasterLogUI')
      .addToUi();
}

// Helper to prevent "null" crashes if script loses context
function getSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Could not connect to Google Sheets. Make sure this script is attached to a Google Sheet.");
  return ss;
}

// This function AUTOMATICALLY computes data when you manually type into the Google Sheet
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  if (sheetName !== "Master Log" && !sheetName.startsWith("Grade ")) return;
  
  const startRow = e.range.getRow();
  if (startRow < 2) return; 
  
  const numRows = e.range.getNumRows();
  const range = sheet.getRange(startRow, 1, numRows, 14);
  const valuesArray = range.getValues();
  
  let updatedAny = false;

  for (let i = 0; i < numRows; i++) {
    const values = valuesArray[i];
    
    let wordCount = parseFloat(values[3]);
    let timeSeconds = parseFloat(values[4]);
    let wpm = parseFloat(values[5]);
    let miscues = parseFloat(values[6]);
    let wordRdgPct = parseFloat(values[7]);
    let wordRdgProfile = values[8];
    let compPct = parseFloat(values[9]);
    let compProfile = values[10];
    let overallProfile = values[11];
    
    let rowUpdated = false;

    if (!isNaN(wordCount) && wordCount > 0 && !isNaN(timeSeconds) && timeSeconds > 0) {
      let calcWpm = (wordCount / timeSeconds) * 60;
      if (isNaN(wpm) || Math.abs(wpm - calcWpm) > 0.1) { values[5] = calcWpm.toFixed(1); rowUpdated = true; }
    }
    if (!isNaN(wordCount) && wordCount > 0 && !isNaN(miscues)) {
      let calcWordPct = ((wordCount - miscues) / wordCount) * 100;
      if (isNaN(wordRdgPct) || Math.abs(wordRdgPct - calcWordPct) > 0.1) { values[7] = calcWordPct.toFixed(1); wordRdgPct = calcWordPct; rowUpdated = true; }
    }
    if (!isNaN(wordRdgPct)) {
      let calcWordProf = getWordReadingLevel(Math.round(wordRdgPct));
      if (wordRdgProfile !== calcWordProf) { values[8] = calcWordProf; wordRdgProfile = calcWordProf; rowUpdated = true; }
    }
    if (!isNaN(compPct)) {
      let calcCompProf = getComprehensionLevel(Math.round(compPct));
      if (compProfile !== calcCompProf) { values[10] = calcCompProf; compProfile = calcCompProf; rowUpdated = true; }
    }
    if (wordRdgProfile && compProfile) {
      let calcOverall = getOverallProfile(wordRdgProfile, compProfile);
      if (overallProfile !== calcOverall) { values[11] = calcOverall; rowUpdated = true; }
    }
    if (rowUpdated) { valuesArray[i] = values; updatedAny = true; }
  }

  if (updatedAny) range.setValues(valuesArray);
}


function initCheck() {
  const ss = getSpreadsheet();
  if (!ss.getSheetByName('Master Log')) setupSpreadsheet();
}

function setupSpreadsheet() {
  const lock = LockService.getScriptLock();
  if (lock && !lock.tryLock(10000)) return; // Prevents crash if running twice
  
  try {
    const ss = getSpreadsheet();

    const FULL_HEADERS = [
      'Date', 'Student Name', 'Grade Level', 'Passage Word Count', 'Time (Seconds)', 
      'Reading Rate (WPM)', 'Miscues (Errors)', 'Word Reading %', 'Word Reading Profile',
      'Comprehension %', 'Comprehension Profile', 'Overall Profile', 'Incorrect Words', 'Miscue Breakdown'
    ];

    let logSheet = ss.getSheetByName('Master Log');
    if (!logSheet) {
      logSheet = ss.insertSheet('Master Log');
      logSheet.appendRow(FULL_HEADERS);
      logSheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff").setHorizontalAlignment("center");
      logSheet.setFrozenRows(1); logSheet.setColumnWidths(1, 14, 150);
    } else if (logSheet.getLastColumn() < 14) {
      logSheet.getRange(1, 1, 1, 14).setValues([FULL_HEADERS]);
      logSheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff").setHorizontalAlignment("center");
      logSheet.setColumnWidths(1, 14, 150);
    }

    let rmSheet = ss.getSheetByName('Reading Material');
    if (!rmSheet) {
      rmSheet = ss.insertSheet('Reading Material');
      rmSheet.appendRow(['Grade Level', 'Date Configured', 'Word Count', 'Total Questions']);
      rmSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff").setHorizontalAlignment("center");
      rmSheet.setFrozenRows(1); rmSheet.setColumnWidths(1, 4, 150);
    }

    let studentsSheet = ss.getSheetByName('Enrolled Students');
    if (!studentsSheet) {
      studentsSheet = ss.insertSheet('Enrolled Students');
      studentsSheet.appendRow(['Student Name', 'Grade Level', 'Enrollment Date']);
      studentsSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1e40af").setFontColor("#ffffff").setHorizontalAlignment("center");
      studentsSheet.setFrozenRows(1); studentsSheet.setColumnWidths(1, 3, 200);
      let sheet1 = ss.getSheetByName('Sheet1');
      if (sheet1) ss.deleteSheet(sheet1);
    }

    let pendingSheet = ss.getSheetByName('Pending Assessments');
    if (!pendingSheet) {
      pendingSheet = ss.insertSheet('Pending Assessments');
      pendingSheet.appendRow(['ID', 'Data JSON Payload']);
      pendingSheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#b91c1c").setFontColor("#ffffff").setHorizontalAlignment("center");
      pendingSheet.setFrozenRows(1); pendingSheet.setColumnWidth(1, 150); pendingSheet.setColumnWidth(2, 600);
    }
  } finally {
    if (lock) lock.releaseLock();
  }
}

function resetSystem(password) {
  if (password !== "03201997") throw new Error("Unauthorized access.");
  const ss = getSpreadsheet();
  ['Master Log', 'Reading Material', 'Enrolled Students', 'Pending Assessments'].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet && sheet.getMaxRows() > 1) sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clearContent();
  });
  ss.getSheets().forEach(sheet => { if (sheet.getName().startsWith('Grade ')) ss.deleteSheet(sheet); });
  return true;
}

function syncMasterLogUI() {
  const ui = SpreadsheetApp.getUi();
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master Log');
  if (!logSheet) { ui.alert("Error: Master Log tab not found. Please run Setup first."); return; }
  if (logSheet.getLastRow() <= 1) { ui.alert("The Master Log is currently empty. There is nothing to sync."); return; }

  const response = ui.alert(
    'Forced Sync Validation', 
    'This will completely WIPE all Grade tabs to make them an exact, deduplicated mirror of the Master Log.\n\nDo you want to proceed?', 
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) {
    const syncCount = coreSyncLogic();
    ui.alert(`Forced Sync Complete!\n\nSuccessfully mirrored ${syncCount} unique records. All Grade tabs are now perfectly synced.`);
  }
}

function coreSyncLogic() {
  const ss = getSpreadsheet();
  const logSheet = ss.getSheetByName('Master Log');
  if (!logSheet) return 0;
  
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) return 0;

  const logData = logSheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const uniqueDataMap = new Map();
  
  logData.forEach(row => {
    const name = String(row[1]).trim().toLowerCase();
    if (!name) return;
    
    // UNBREAKABLE FINGERPRINT: Ignores date completely. Matches Name + Grade + WordCount + Time
    const grade = String(row[2]).trim();
    const wordCount = String(row[3]).trim();
    const time = String(row[4]).trim();
    const signature = `${name}_${grade}_${wordCount}_${time}`;
    
    if (uniqueDataMap.has(signature)) {
       const existingRow = uniqueDataMap.get(signature);
       // Keep the one that is most "complete" (has an Overall Profile)
       if (String(row[11]).trim() !== "" && String(existingRow[11]).trim() === "") {
           uniqueDataMap.set(signature, row);
       }
    } else {
      uniqueDataMap.set(signature, row);
    }
  });

  const uniqueData = Array.from(uniqueDataMap.values());

  // Write cleaned data back to Master Log
  const maxLogRows = logSheet.getMaxRows();
  if (maxLogRows > 1) logSheet.getRange(2, 1, maxLogRows - 1, 14).clearContent();
  if (uniqueData.length > 0) logSheet.getRange(2, 1, uniqueData.length, 14).setValues(uniqueData);

  const gradeGroups = {};
  uniqueData.forEach(row => {
    const grade = String(row[2]).trim();
    if (grade && grade !== "Unknown") {
      if (!gradeGroups[grade]) gradeGroups[grade] = [];
      gradeGroups[grade].push(row);
    }
  });

  const sheets = ss.getSheets();
  let syncCount = 0;

  // STRICT WIPING: Clear all Grade tabs regardless of filters (ignores visibility)
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('Grade ')) {
      const gradeStr = name.replace('Grade ', '').trim();
      const maxRows = sheet.getMaxRows();
      if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, Math.max(14, sheet.getMaxColumns())).clearContent();
      
      if (gradeGroups[gradeStr] && gradeGroups[gradeStr].length > 0) {
         sheet.getRange(2, 1, gradeGroups[gradeStr].length, 14).setValues(gradeGroups[gradeStr]);
         syncCount += gradeGroups[gradeStr].length;
         delete gradeGroups[gradeStr];
      }
    }
  });

  for (const grade in gradeGroups) {
     if (gradeGroups[grade].length > 0) {
       addGradeLevel(grade); 
       const newSheet = ss.getSheetByName(`Grade ${grade}`);
       newSheet.getRange(2, 1, gradeGroups[grade].length, 14).setValues(gradeGroups[grade]);
       syncCount += gradeGroups[grade].length;
     }
  }
  return syncCount;
}

function getAvailableGrades() {
  initCheck(); 
  const ss = getSpreadsheet();
  const sheets = ss.getSheets();
  const gradesSet = new Set();

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name.startsWith('Grade ')) {
      const gradeStr = name.replace('Grade ', '').trim();
      if (gradeStr) gradesSet.add(gradeStr);
    }
  });

  const studentsSheet = ss.getSheetByName('Enrolled Students');
  if (studentsSheet && studentsSheet.getLastRow() > 1) {
     const data = studentsSheet.getRange(2, 2, studentsSheet.getLastRow() - 1, 1).getValues();
     data.forEach(row => { const g = String(row[0]).trim(); if (g) gradesSet.add(g); });
  }

  const grades = Array.from(gradesSet);
  return grades.sort((a, b) => {
    const numA = parseInt(a); const numB = parseInt(b);
    if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
    return a.localeCompare(b);
  });
}

function addGradeLevel(grade) {
  if (!grade) throw new Error("Grade is required.");
  const ss = getSpreadsheet();
  const tabName = `Grade ${grade}`;
  let gradeSheet = ss.getSheetByName(tabName);
  if (gradeSheet) throw new Error(`Grade ${grade} already exists.`);

  gradeSheet = ss.insertSheet(tabName);
  gradeSheet.appendRow([
    'Date', 'Student Name', 'Grade Level', 'Passage Word Count', 'Time (Seconds)',
    'Reading Rate (WPM)', 'Miscues (Errors)', 'Word Reading %', 'Word Reading Profile',
    'Comprehension %', 'Comprehension Profile', 'Overall Profile', 'Incorrect Words', 'Miscue Breakdown'
  ]);
  gradeSheet.getRange(1, 1, 1, 14).setFontWeight("bold").setBackground("#15803d").setFontColor("#ffffff").setHorizontalAlignment("center").setVerticalAlignment("middle");
  gradeSheet.setFrozenRows(1); gradeSheet.setColumnWidths(1, 14, 150);

  return getAvailableGrades();
}

function getGradeData(gradeLevel) {
  const ss = getSpreadsheet();
  const studentsSheet = ss.getSheetByName('Enrolled Students');
  let students = [];
  
  if (studentsSheet && studentsSheet.getLastRow() > 1) {
    const data = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 2).getValues();
    students = data
      .filter(row => String(row[1]).trim() === String(gradeLevel).trim() && String(row[0]).trim() !== "")
      .map(row => ({ name: String(row[0]).trim(), grade: String(row[1]).trim() }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }

  const rmSheet = ss.getSheetByName('Reading Material');
  let settings = { wordCount: '', compTotal: '' };
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  if (rmSheet && rmSheet.getLastRow() > 1) {
    const rmData = rmSheet.getRange(2, 1, rmSheet.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < rmData.length; i++) {
      if (String(rmData[i][0]).trim() === String(gradeLevel).trim()) {
        let rowDate = rmData[i][1];
        if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
        else { try { rowDate = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), "MM/dd/yyyy"); } catch(e) { rowDate = String(rmData[i][1]); } }
        
        if (rowDate === currentDate) { settings.wordCount = rmData[i][2] || ''; settings.compTotal = rmData[i][3] || ''; }
        break; 
      }
    }
  }

  return { students: students, settings: settings };
}

function saveDailySettings(gradeLevel, wordCount, compTotal) {
  const ss = getSpreadsheet();
  const rmSheet = ss.getSheetByName('Reading Material');
  if (!rmSheet) return false;

  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  const lastRow = rmSheet.getLastRow();
  let foundRow = -1;

  if (lastRow > 1) {
    const rmData = rmSheet.getRange(2, 1, lastRow - 1, 1).getValues(); 
    for (let i = 0; i < rmData.length; i++) {
      if (String(rmData[i][0]).trim() === String(gradeLevel).trim()) { foundRow = i + 2; break; }
    }
  }

  if (foundRow !== -1) { rmSheet.getRange(foundRow, 2).setValue(currentDate); rmSheet.getRange(foundRow, 3).setValue(wordCount); rmSheet.getRange(foundRow, 4).setValue(compTotal); } 
  else rmSheet.appendRow([gradeLevel, currentDate, wordCount, compTotal]);
  
  return { wordCount: wordCount, compTotal: compTotal };
}

function searchAllStudents(query) {
  const ss = getSpreadsheet();
  const studentsSheet = ss.getSheetByName('Enrolled Students');
  if (!studentsSheet || studentsSheet.getLastRow() <= 1) return [];
  
  const data = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 2).getValues();
  const q = query.toLowerCase();
  
  return data
    .filter(row => row[0].toString().toLowerCase().includes(q))
    .map(row => ({ name: row[0], grade: row[1] }))
    .sort((a, b) => a.name.localeCompare(b.name))
    .slice(0, 10); 
}

function enrollStudentsBulk(studentNames, gradeLevel) {
  if (!Array.isArray(studentNames) || studentNames.length === 0 || !gradeLevel) throw new Error("Student Names array and Grade Level are required.");
  const ss = getSpreadsheet();

  const studentsSheet = ss.getSheetByName('Enrolled Students');
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  let existingNames = [];
  if (studentsSheet.getLastRow() > 1) existingNames = studentsSheet.getRange(2, 1, studentsSheet.getLastRow() - 1, 1).getValues().flat();

  const addedStudents = [];
  let gradeSheet = ss.getSheetByName(`Grade ${gradeLevel}`);
  if (!gradeSheet) addGradeLevel(gradeLevel);

  studentNames.forEach(name => {
    const trimmedName = name.trim();
    if (trimmedName && !existingNames.includes(trimmedName)) {
      studentsSheet.appendRow([trimmedName, gradeLevel, currentDate]);
      addedStudents.push(trimmedName);
    }
  });

  return { added: addedStudents };
}

function upsertPendingAssessment(data) {
  const ss = getSpreadsheet();
  const pendingSheet = ss.getSheetByName('Pending Assessments');
  if (!pendingSheet) return false;

  let rowToUpdate = -1;
  if (pendingSheet.getLastRow() > 1) {
    const ids = pendingSheet.getRange(2, 1, pendingSheet.getLastRow() - 1, 1).getValues();
    for(let i = 0; i < ids.length; i++) {
      if(ids[i][0].toString() === data.id.toString()) { rowToUpdate = i + 2; break; }
    }
  }

  const payloadStr = JSON.stringify(data);
  if (rowToUpdate !== -1) pendingSheet.getRange(rowToUpdate, 2).setValue(payloadStr);
  else pendingSheet.appendRow([data.id, payloadStr]);
  
  return true;
}

function getPendingAssessments() {
  const ss = getSpreadsheet();
  const pendingSheet = ss.getSheetByName('Pending Assessments');
  if (!pendingSheet || pendingSheet.getLastRow() <= 1) return [];

  const data = pendingSheet.getRange(2, 1, pendingSheet.getLastRow() - 1, 2).getValues();
  return data.map(row => { try { return JSON.parse(row[1]); } catch(e) { return null; } }).filter(item => item !== null);
}

function deletePendingAssessment(id) {
  const ss = getSpreadsheet();
  const pendingSheet = ss.getSheetByName('Pending Assessments');
  if (!pendingSheet || pendingSheet.getLastRow() <= 1) return;

  const data = pendingSheet.getRange(2, 1, pendingSheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) { pendingSheet.deleteRow(i + 2); break; }
  }
}

function saveAssessmentData(data) {
  const ss = getSpreadsheet();
  const rawDate = new Date();
  const dateString = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
  const gradeLevel = data.grade || "Unknown";

  const logSheet = ss.getSheetByName('Master Log');
  if (!logSheet) throw new Error("Could not find a tab named 'Master Log'");
  
  const incorrectWordsStr = (data.incorrectWords && Array.isArray(data.incorrectWords)) ? data.incorrectWords.join(', ') : "";
  const miscueBreakdownStr = data.miscueBreakdown ? Object.entries(data.miscueBreakdown).filter(([k,v]) => v > 0).map(([k,v]) => `${k.charAt(0).toUpperCase() + k.slice(1)}: ${v}`).join(', ') : "";
  
  const rowData = [
    dateString, data.student, gradeLevel, data.wordCount, data.timeInSeconds, 
    data.wpm, data.miscues, data.wordReadingPercent, data.wordReadingProfile, 
    data.comprehensionPercent, data.comprehensionProfile, data.finalProfile, 
    incorrectWordsStr, miscueBreakdownStr
  ];

  let isUpdate = false;

  if (data.originalFingerprint) {
     const logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 14).getValues();
     for (let i = 0; i < logData.length; i++) {
         const rowName = String(logData[i][1]).trim().toLowerCase();
         const rowGrade = String(logData[i][2]).trim();
         const rowWc = String(logData[i][3]).trim();
         const rowTime = String(logData[i][4]).trim();
         const sig = `${rowName}_${rowGrade}_${rowWc}_${rowTime}`;
         
         if (sig === data.originalFingerprint) {
             logSheet.getRange(i + 2, 1, 1, 14).setValues([rowData]);
             isUpdate = true; break;
         }
     }
  }

  if (!isUpdate) logSheet.appendRow(rowData);
  if (gradeLevel !== "Unknown") saveDailySettings(gradeLevel, data.wordCount, data.compTotal);

  coreSyncLogic();
  if (data.id) deletePendingAssessment(data.id);
  return "Success";
}

function getResultsForPDF(gradeLevel, startDate, endDate) {
  const ss = getSpreadsheet();
  const logSheet = ss.getSheetByName('Master Log');
  if (!logSheet || logSheet.getLastRow() <= 1) return [];

  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 14).getValues();
  const start = new Date(startDate); start.setHours(0, 0, 0, 0);
  const end = new Date(endDate); end.setHours(23, 59, 59, 999);
  const results = [];

  for (let i = 0; i < data.length; i++) {
    if (data[i][2].toString() === gradeLevel.toString()) {
      let rowDate = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
      if (rowDate >= start && rowDate <= end) {
        results.push({
          date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "MM/dd/yyyy"),
          studentName: data[i][1], wpm: data[i][5], miscues: data[i][6], 
          wordScore: data[i][7], wordProfile: data[i][8], compScore: data[i][9], 
          compProfile: data[i][10], profile: data[i][11], incorrectWords: data[i][12] || ""
        });
      }
    }
  }
  return results;
}

function getStudentHistory(studentName) {
  const ss = getSpreadsheet();
  const logSheet = ss.getSheetByName('Master Log');
  if (!logSheet || logSheet.getLastRow() <= 1) return []; 
  
  const data = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 14).getValues();
  const recentRecords = data.filter(row => String(row[1]).trim().toLowerCase() === String(studentName).trim().toLowerCase()).slice(-3).reverse();
  
  return recentRecords.map(row => {
    let dateStr = "Unknown Date";
    try {
        let dateVal = row[0];
        if (!(dateVal instanceof Date)) dateVal = new Date(dateVal);
        if (!isNaN(dateVal.getTime())) dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "MM/dd/yyyy");
        else dateStr = String(row[0]); 
    } catch(e) { dateStr = String(row[0]); }

    return {
      date: dateStr, wpm: row[5], wordScore: row[7], wordProfile: row[8],
      compScore: row[9], compProfile: row[10], profile: row[11],
      gradeLevel: row[2], wordCount: row[3], timeInSeconds: row[4], miscueCount: row[6],
      incorrectWords: row[12] ? String(row[12]).split(', ') : []
    };
  });
}
