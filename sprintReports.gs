function createSprintReport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Active Sprint Data B');
  var targetSheetName = 'SprintReport';

  // Delete the target sheet if it already exists
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }

  // Create the target sheet
  targetSheet = ss.insertSheet(targetSheetName);

  // Get the data from the source sheet
  var data = sourceSheet.getDataRange().getValues();

  // Find the index of the relevant columns
  var headers = data[0];
  var assigneeIndex = headers.indexOf('Assignee');
  var sprintStateIndex = headers.indexOf('Sprint State');
  var storyPointsIndex = headers.indexOf('Story Points');
  var sprintStartDateIndex = headers.indexOf('Sprint Start Date');

  // Create a unique list of assignees and initialize tables
  var assignees = {};
  for (var i = 1; i < data.length; i++) {
    var assignee = data[i][assigneeIndex];
    if (assignee && !assignees[assignee]) {
      assignees[assignee] = [];
    }
    assignees[assignee].push(data[i]);
  }

  // Create tables for each assignee
  var row = 1;
  for (var assignee in assignees) {
    // Set Assignee name as a header
    targetSheet.getRange(row, 1, 1, 4).merge();
    targetSheet.getRange(row, 1).setValue(assignee)
      .setFontWeight('bold')
      .setFontSize(12)
      .setHorizontalAlignment('center')
      .setBackground('#e91f64')
      .setFontColor('#fef1f7');

    // Set headers 'ID', 'Active', 'Closed', 'Difference'
    targetSheet.getRange(row + 1, 1).setValue('ID'); // Set 'ID' as a column header
    targetSheet.getRange(row + 1, 2).setValue('Active'); // Set 'Active' as a column header
    targetSheet.getRange(row + 1, 3).setValue('Closed'); // Set 'Closed' as a column header
    targetSheet.getRange(row + 1, 4).setValue('Difference'); // Set 'Difference' as a column header

    // Make headers bold and apply specified formatting
    targetSheet.getRange(row + 1, 1, 1, 4)
      .setFontWeight('bold')
      .setBackground('#fecce3')
      .setFontColor('#55021a')
      .setHorizontalAlignment('center');

    var assigneeData = assignees[assignee];
    var uniqueDates = [...new Set(assigneeData.map(item => new Date(item[sprintStartDateIndex]).toISOString().split('T')[0]))];

    var currentRow = row + 2;
    uniqueDates.forEach(date => {
      var activePoints = 0;
      var closedPoints = 'Pending';

      assigneeData.forEach(item => {
        if (new Date(item[sprintStartDateIndex]).toISOString().split('T')[0] === date) {
          if (item[sprintStateIndex] === 'active') {
            activePoints += item[storyPointsIndex] || 0;
          } else if (item[sprintStateIndex] === 'closed') {
            if (closedPoints === 'Pending') {
              closedPoints = 0;
            }
            closedPoints += item[storyPointsIndex] || 0;
          }
        }
      });

      var differencePoints = closedPoints === 'Pending' ? 'Pending' : closedPoints - activePoints;

      targetSheet.getRange(currentRow, 1).setValue(date); // Set Date
      if (closedPoints === 'Pending') {
        targetSheet.getRange(currentRow, 2).setValue(activePoints); // Set Active Points only if sprint state is active
      }
      targetSheet.getRange(currentRow, 3).setValue(closedPoints); // Set Closed Points
      targetSheet.getRange(currentRow, 4).setValue(differencePoints); // Set Difference Points
      currentRow++;
    });

    // Apply borders around the table
    targetSheet.getRange(row, 1, currentRow - row, 4).setBorder(true, true, true, true, true, true);

    row = currentRow + 1; // Move to the next header position
  }
}