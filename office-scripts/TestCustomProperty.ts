/**
 * TestCustomProperty – Office Script to test co-authoring sync
 *
 * Writes a custom document property with a timestamp.
 * Run this from Power Automate, then check if the add-in
 * can see the updated value without a page refresh.
 */
function main(workbook: ExcelScript.Workbook): string {
  const timestamp = new Date().toISOString();
  const command = JSON.stringify({
    type: "SRK_HIGHLIGHT_STUDENT_ROW",
    data: {
      syStudentId: "10003",
      targetSheet: "Master List",
      startCol: "Student Name",
      endCol: "Outreach",
      color: "#FFFF00"
    },
    timestamp: timestamp
  });

  // Write to custom document property
  workbook.getProperties().addCustomProperty("SRK_Command", command);

  return `SUCCESS: Custom property SRK_Command set at ${timestamp}`;
}
