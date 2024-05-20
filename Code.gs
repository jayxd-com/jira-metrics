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

    var row = [
      projectName, parentTicket, subTaskTicket, summary, assignee,
      storyPoints, subTaskStatus, parentStatus, sprintName, scrumTeam,
      sprintState, sprintStartDate, sprintEndDate, parentSelf
    ];

    data.push(row);
  });

  // Sort data by Sprint State ASC and then by Assignee ASC
  data.sort(function(a, b) {
    if (a[10] === b[10]) {
      return a[4].localeCompare(b[4]);
    }
    return a[10].localeCompare(b[10]);
  });

  // Filter rows by date
  var filteredData = filterRowsByDate(data);

  // Append filtered data to the sheet
  filteredData.forEach(function(row) {
    sheet.appendRow(row);
  });

  // Create Sprint Reports
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


// THIS BLOCK HAS ISSUES. 
function createSprintReport(data) {
  var sprintReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SprintReports');
  if (!sprintReportSheet) {
    sprintReportSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('SprintReports');
  } else {
    sprintReportSheet.clear(); // clear sheet
  }

  var assigneeData = {};

  data.forEach(function(row) {
    var assignee = row[4];
    var sprintStartDate = row[11] ? new Date(row[11]).toLocaleDateString() : '';
    var storyPoints = row[5];
    var subTaskStatus = row[6];
    var sprintState = row[10];

    if (!assigneeData[assignee]) {
      assigneeData[assignee] = {};
    }

    if (!assigneeData[assignee][sprintStartDate]) {
      assigneeData[assignee][sprintStartDate] = {
        totalSprintPoints: 0,
        totalSprintsNotDone: 0
      };
    }

    if (sprintState === 'active') {
      assigneeData[assignee][sprintStartDate].totalSprintPoints += storyPoints || 0;
    }

    if (subTaskStatus !== 'Done') {
      assigneeData[assignee][sprintStartDate].totalSprintsNotDone += storyPoints || 0;
    }
  });

  var rowIndex = 1;

  for (var assignee in assigneeData) {
    sprintReportSheet.appendRow([`Assignee: ${assignee}`]);
    sprintReportSheet.getRange(rowIndex, 1).setFontWeight("bold");
    rowIndex++;

    var headers = ['ID (Sprint Start Date)', 'Total Sprint Points', 'Total Sprints Not Done'];
    sprintReportSheet.appendRow(headers);
    rowIndex++;

    for (var sprintStartDate in assigneeData[assignee]) {
      var data = assigneeData[assignee][sprintStartDate];
      sprintReportSheet.appendRow([
        sprintStartDate, data.totalSprintPoints, data.totalSprintsNotDone
      ]);
      rowIndex++;
    }
    rowIndex++;
  }
}

function formatDates() {
  // Example date
  var exampleDate = new Date();

  // Format date using Moment.js
  var formattedDate = moment(exampleDate).format('MMMM Do YYYY, h:mm:ss a');

  Logger.log('Formatted Date: ' + formattedDate);
}
