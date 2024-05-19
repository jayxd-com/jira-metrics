function fetchJiraData() {

  formatDates();

  var
    jiraDomain = 'https://moveinc.atlassian.net';

  var
    username = 'mehvish.hashmi.contractor@realtor.com';

  var
    apiToken = 'YOUR_API';

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
    'Issue Key', 'Sub Task Ticket', 'Parent Self', 'Project Name', 'Summary', 'Assignee',
    'Issue Type', 'Story Points', 'Actual Story Points+', 'Status', 'Parent Status', 'Sprint Name',
    'Sprint State', 'Sprint Start Date', 'Sprint End Date', 'Sprint Start (Human)', 'Sprint End (Human)',
    'Sprint Complete Date', 'Sprint Complete Date (Human)'
  ];

  sheet.appendRow(headers);

  // Styling headers - bold and increase font size
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setFontSize(12);

  var assigneeMap = {};
  var projectSprintsMap = {};

  issues.forEach(function(issue) {
    var issueKey = issue.key;
    var parentSelf = issue.fields.parent ? issue.fields.parent.self : '';
    var subTaskTicket = `https://moveinc.atlassian.net/browse/${issueKey}`;
    var projectName = issue.fields.project.name;
    var summary = issue.fields.summary;
    var assignee = issue.fields.assignee ? issue.fields.assignee.displayName : 'Unassigned';
    var issueType = issue.fields.issuetype.name;
    var storyPoints = issue.fields['customfield_10034']; // Story Points for all issues
    var actualStoryPoints = issue.fields.parent ? issue.fields.parent.fields['customfield_10548'] : '';
    var status = issue.fields.status.name;
    var parentStatus = issue.fields.parent ? issue.fields.parent.fields.status.name : '';
    var sprint = issue.fields['customfield_10020'] ? issue.fields['customfield_10020'][0] : {};
    var sprintName = sprint.name || '';
    var sprintState = sprint.state || '';
    var sprintStartDate = sprint.startDate || '';
    var sprintEndDate = sprint.endDate || '';
    var sprintCompleteDate = sprint.completeDate || '';

    // Format dates in human-readable format using Moment.js
    var sprintStartHuman = sprintStartDate ? moment(sprintStartDate).fromNow() : '';
    var sprintEndHuman = sprintEndDate ? moment(sprintEndDate).fromNow() : '';
    var sprintCompleteHuman = sprintCompleteDate ? moment(sprintCompleteDate).fromNow() : '';

    if (assignee !== 'Unassigned') {
      assigneeMap[assignee] = assignee;
    }

    var row = [
      issueKey, subTaskTicket, parentSelf, projectName, summary, assignee, issueType,
      storyPoints, actualStoryPoints, status, parentStatus, sprintName, sprintState,
      sprintStartDate, sprintEndDate, sprintStartHuman, sprintEndHuman,
      sprintCompleteDate, sprintCompleteHuman
    ];
    sheet.appendRow(row);

    // Group sprints by project
    if (!projectSprintsMap[projectName]) {
      projectSprintsMap[projectName] = { active: [], closed: [] };
    }
    if (sprintState === 'active') {
      projectSprintsMap[projectName].active.push(row);
    } else if (sprintState === 'closed') {
      projectSprintsMap[projectName].closed.push(row);
    }
  });

  // Create a new sheet for active sprints and the last closed sprint per project
  var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Active Sprint Data 2');
  if (!sheet2) {
    sheet2 = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Active Sprint Data 2');
  } else {
    sheet2.clear(); // clear sheet
  }

  sheet2.appendRow(headers);
  headerRange = sheet2.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setFontSize(12);

  // Add active sprints and the last closed sprint per project to the new sheet
  for (var project in projectSprintsMap) {
    var sprints = projectSprintsMap[project];
    sprints.active.forEach(function(row) {
      sheet2.appendRow(row);
    });
    if (sprints.closed.length > 0) {
      var lastClosedSprint = sprints.closed[sprints.closed.length - 1];
      sheet2.appendRow(lastClosedSprint);
    }
  }


}

  var velocitySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Velocity');
  if (!velocitySheet) {
    velocitySheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Velocity');
  } else {
    velocitySheet.clear(); // This will clear sheet data.
  }

  var headers = ['Assignee', 'Sprint', 'Initial Story Points', 'Completed Story Points', 'Rolled Over Story Points'];
  velocitySheet.appendRow(headers);

  // Styling Header - bold and increase font size
  var headerRange = velocitySheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold").setFontSize(12);

  var assigneePointsMap = {};

  for (var project in projectSprintsMap) {
    var sprints = projectSprintsMap[project];

    // Iterate through active sprints
    sprints.active.forEach(function(row) {
      var assignee = row[5];
      var storyPoints = row[7] || row[8];
      var sprintName = row[11];
      var status = row[9];

      if (!assigneePointsMap[assignee]) {
        assigneePointsMap[assignee] = {};
      }
      if (!assigneePointsMap[assignee][sprintName]) {
        assigneePointsMap[assignee][sprintName] = {
          initial: 0,
          completed: 0,
          rolledOver: 0
        };
      }

      assigneePointsMap[assignee][sprintName].initial += storyPoints || 0;
      if (status === 'Done') {
        assigneePointsMap[assignee][sprintName].completed += storyPoints || 0;
      } else {
        assigneePointsMap[assignee][sprintName].rolledOver += storyPoints || 0;
      }
    });

    // Handle the last closed sprint
    if (sprints.closed.length > 0) {
      var lastClosedSprint = sprints.closed[sprints.closed.length - 1];
      var assignee = lastClosedSprint[5];
      var storyPoints = lastClosedSprint[7] || lastClosedSprint[8];
      var sprintName = lastClosedSprint[11];
      var status = lastClosedSprint[9];

      if (!assigneePointsMap[assignee]) {
        assigneePointsMap[assignee] = {};
      }
      if (!assigneePointsMap[assignee][sprintName]) {
        assigneePointsMap[assignee][sprintName] = {
          initial: 0,
          completed: 0,
          rolledOver: 0
        };
      }

      assigneePointsMap[assignee][sprintName].completed += storyPoints || 0;
    }
  }

  // Write data to the Velocity sheet
  for (var assignee in assigneePointsMap) {
    for (var sprint in assigneePointsMap[assignee]) {
      var sprintData = assigneePointsMap[assignee][sprint];
      velocitySheet.appendRow([assignee, sprint, sprintData.initial, sprintData.completed, sprintData.rolledOver]);
    }
  }