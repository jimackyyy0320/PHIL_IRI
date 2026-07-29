function getWordReadingLevel(score) {
  if (score >= 97) return "Independent";
  if (score >= 90) return "Instructional";
  return "Frustration";
}

function getComprehensionLevel(score) {
  if (score >= 80) return "Independent";
  if (score >= 59) return "Instructional";
  return "Frustration";
}

function getOverallProfile(wrLevel, compLevel) {
  if (wrLevel === "Independent" && compLevel === "Independent") return "Independent";
  if (wrLevel === "Independent" && compLevel === "Instructional") return "Instructional";
  if (wrLevel === "Instructional" && compLevel === "Independent") return "Instructional";
  if (wrLevel === "Instructional" && compLevel === "Instructional") return "Instructional";
  if (wrLevel === "Instructional" && compLevel === "Frustration") return "Frustration";
  if (wrLevel === "Frustration" && compLevel === "Instructional") return "Frustration";
  if (wrLevel === "Frustration" && compLevel === "Frustration") return "Frustration";
  return "Frustration";
}

function getSharedFunctions() {
  return getWordReadingLevel.toString() + "\n" +
         getComprehensionLevel.toString() + "\n" +
         getOverallProfile.toString();
}
