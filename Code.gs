function fetchJiraData() {
  formatDates();

  var jiraDomain = 'https://moveinc.atlassian.net';
  var username = 'mehvish.hashmi.contractor@realtor.com';
  var apiToken = 'YOUR_API_TOKEN';

  var jqlQuery = "project in (MTECH, ANDROID, IOS, 'TL-TM Android', 'TL-TM iOS', 'TL-TM MTECH') AND sprint in openSprints() AND issueType = sub-task AND assignee IN (712020:4bb113e3-a8bc-4735-81b2-65e06022ccfe, 712020:c2155c7e-5b33-4b53-a4e2-5c763506a4f4, " +
    "712020:48ab14ca-5ecf-427c-9324-953f28b1e639, 712020:09939799-9b85-4101-afc9-88f749b195b7, 712020:03d0778c-ea5a-4b9e-a845-e4bbf4315bc0, 712020:f591dbe6-1a63-4dbd-9aef-d7af539d9d20, 712020:4bb113e3-a8bc-4735-81b2-65e06022ccfe, 62df5698f6dd8b8b0eab6dc8, 712020:c4c0de3c-c401-46d9-b4da-d9c37bf32236)";

  var encodedJql = encodeURIComponent(jqlQuery);
  var url = `${jiraDomain}/rest/api/2/search?jql=${encodedJql}&expand=names,fields`;

  var headers = {
    "Authorization": "Basic " + Utilities.base64Encode(username + ':' + apiToken),
    "Accept": "application/json"
  };

  var options = {
    "method": "GET",
    "headers": headers
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());
  var issues = json.issues;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Active Sprint Data');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Active Sprint Data');
  } else {
    sheet.clear(); // clear sheet
  }

  // You can update the header or rearrange headers from here.

  var headers = [
    'Project Name', 'Parent Ticket', 'Sub Task Ticket', 'Summary', 'Assignee',
    'Story Points', 'SubTask Status', 'Parent Status', 'Sprint Name', 'Scrum Team',
    'Sprint State', 'Sprint Start Date', 'Sprint End Date', 'Parent Self'
  ];

  sheet.appendRow(headers);

  // Styling headers - bold and increase font size
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setFontSize(12);

  var data = [];

  // If you want to add or remove any columns, you can right here. Set a variable 'var' and assign value. These values are coming from the API. you can use postman to view structure. 
  // For example, issue.fields.parent means item under issues having the key of fields and then parent. Its an array. 
// If you want to target, my_example key that is under issues > fields > parent, you will type 
// var example = issue.fields.parent.my_example; 
// Adding ? apply if() {} else {}  
// (condition) ? 'this' or 'that';

  issues.forEach(function(issue) {
    var parentTicket = issue.fields.parent ? `https://moveinc.atlassian.net/browse/${issue.fields.parent.key}` : '';
    var parentSelf = issue.fields.parent ? issue.fields.parent.self : '';
    var subTaskTicket = `https://moveinc.atlassian.net/browse/${issue.key}`;
    var projectName = issue.fields.project.name;
    var summary = issue.fields.summary;
    var assignee = issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned';
    var storyPoints = issue.fields['customfield_10034']; // Story Points for all issues
    var subTaskStatus = issue.fields.status.name;
    var parentStatus = issue.fields.parent ? issue.fields.parent.fields.status.name : '';
    var sprint = issue.fields['customfield_10020'] ? issue.fields['customfield_10020'][0] : {};
    var sprintName = sprint.name || '';
    var sprintState = sprint.state || '';
    var sprintStartDate = sprint.startDate || '';
    var sprintEndDate = sprint.endDate || '';
    var scrumTeam = issue.fields['customfield_10137'] ? issue.fields['customfield_10137'].value : '';


// These are the values for columns. Make sure they are aligned with headers. You can create or remove variables above and use them here. You can rearrange values here if you move header.
    var row = [
      projectName, parentTicket, subTaskTicket, summary, assignee,
      storyPoints, subTaskStatus, parentStatus, sprintName, scrumTeam,
      sprintState, sprintStartDate, sprintEndDate, parentSelf
    ];

// This is a forEach loop. At every iteration, it will pust "row array " to new array data. Later Data will have all rows for the sheet. 
    data.push(row);
  });


    // Here we are sorting the data as mentioned below. If you rearrange columns, make sure to update index too.
  // Sort data by Sprint State ASC and then by Assignee ASC
  data.sort(function(a, b) {
    if (a[10] === b[10]) {
      return a[4].localeCompare(b[4]);
    }
    return a[10].localeCompare(b[10]);
  });

    // Here we are filtering the data so we will see only 2 weeks of data. It will run the function below. 
  // Filter rows by date
  var filteredData = filterRowsByDate(data);

  // Append filtered data to the sheet
  filteredData.forEach(function(row) {
    sheet.appendRow(row);
  });

  // Create Sprint Reports
// ===========================
// NOTE:
// Uncomment this section only when you have applied function from sprintReport here. So it will run. It should always run in last after first sheet is created.


// I haven't testing this section due to API restriction. It should work based on my guess.

  createSprintReport(filteredData);

}

function filterRowsByDate(data) {
  var filteredData = [];
  var activeSprintStartDate;

  // Find the Sprint Start Date where Sprint State is active
  data.forEach(function(row) {
    if (row[10] === 'active') {
      activeSprintStartDate = row[11];
    }
  });

  // Filter rows to exclude Sprint State closed and date minus 15 days
  if (activeSprintStartDate) {
    var targetDate = new Date(activeSprintStartDate);
    targetDate.setDate(targetDate.getDate() - 15);

    data.forEach(function(row) {
      if (row[10] === 'active' || (row[10] === 'closed' && new Date(row[11]) > targetDate)) {
        filteredData.push(row);
      }
    });
  }

  return filteredData;
}


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


// THIS IS FOR TESTING ONLY. IT CAN BE REMOVED.
function formatDates() {
  // Example date
  var exampleDate = new Date();

  // Format date using Moment.js
  var formattedDate = moment(exampleDate).format('MMMM Do YYYY, h:mm:ss a');

  Logger.log('Formatted Date: ' + formattedDate);
}
