function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('JIRA Issue Tracking')
      .addItem('Edit JIRA settings', 'editJIRASettings')
      .addItem('Edit sheet query', 'editJIRAQuery')
      .addItem('Fetch current statuses', 'loadAndPopulateJIRA')
      .addItem('Fetch historic statuses', 'loadHistoricStates')
      .addSeparator()
      .addItem('Publish charts to Drive', 'publishDriveCharts')
      .addItem('Edit Confluence settings', 'editConfluenceSettings')
      .addItem('Publish charts to Confluence', 'updateConfluenceCharts')
      .addToUi();
}

function createTeamSheet(name, templateSheet=null) {
  const ss = templateSheet !== null ? SpreadsheetApp.openById(templateSheet) :
    SpreadsheetApp.getActiveSpreadsheet();
  const newSheet = ss.copy(name);
  newSheet.deleteSheet(newSheet.getSheetByName('Team Sheets'));
  return newSheet;
}

function addTriggersIfNotInstalled() {
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0) {
    ScriptApp.newTrigger('loadAndPopulateJIRA')
        .timeBased()
        .atHour(1)
        .nearMinute(0)
        .everyDays(1)
        .create();
    ScriptApp.newTrigger('addChartDateColumns')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.SATURDAY)
        .atHour(1)
        .nearMinute(30)
        .create();
    ScriptApp.newTrigger('updateConfluenceCharts')
        .timeBased()
        .atHour(2)
        .nearMinute(0)
        .everyDays(1)
        .create();
  }
}

function dateAsString_(dateValue) {
  return Utilities.formatDate(dateValue, 'GMT', 'yyyy-MM-dd');
}

function loadAndPopulateJIRA() {
  const now = new Date();
  const todayStr = dateAsString_(now);
  const dataSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(
    (sheet) => sheet.getName().indexOf('Status: ') === 0
  );
  dataSheets.forEach(destSheet => {
    const tempSheetName = destSheet.getName().replace(/^Status: /, 'Query: ') + ' ' + todayStr;
    let newSheet = SpreadsheetApp.getActive().getSheetByName(tempSheetName);
    if (newSheet === null) {
      newSheet = SpreadsheetApp.getActive().insertSheet(tempSheetName);
      const jql = getJIRAQuery(destSheet);
      if (jql === '') {
        throw 'No query was found for sheet ' + destSheet.getName();
      }
      getJIRAIssues_(jql, newSheet);
    }
    let nextColumn = findGivenDateColumn_(destSheet, now);
    if (nextColumn === -1) {
      nextColumn = destSheet.getMaxColumns() + 1;
      destSheet.getRange(1, nextColumn).setValue(todayStr);
    }
    populateAllItems_(newSheet, destSheet, nextColumn);
    newSheet.hideSheet();
  });
}

function loadHistoricStates() {
  const dataSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(
    (sheet) => sheet.getName().indexOf('Status: ') === 0
  );
  dataSheets.forEach(destSheet => {
    fillSheetIssueHistories(destSheet);
  });
}

function getJIRAUrl() {
  const url = PropertiesService.getScriptProperties().getProperty('jiraUrl') || '';
  return url.replace(/\/+$/g)
}

function getJIRAQuery(querySheet=null) {
  querySheet = querySheet || SpreadsheetApp.getActiveSheet();
  const queryMetadata = getSheetMetadataPropertyByKey(querySheet, 'jiraQuery');
  return queryMetadata !== undefined ? queryMetadata.getValue() : '';
}

function getSheetMetadataPropertyByKey(sheet, keyName) {
  const metadata = sheet.getDeveloperMetadata();
  return metadata.filter(function(metadata) {
    return metadata.getLocation().getLocationType() === SpreadsheetApp.DeveloperMetadataLocationType.SHEET;
  }).find(function(metadata) {
    return metadata.getKey() === keyName;
  });
}

function editJIRASettings() {
  let html = HtmlService.createTemplateFromFile('JIRASettings')
      .evaluate()
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit JIRA Settings');
}

function editJIRAQuery() {
  let html = HtmlService.createTemplateFromFile('JIRAQuery')
      .evaluate()
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit JIRA Query');
}

function editConfluenceSettings() {
  let html = HtmlService.createTemplateFromFile('ConfluenceSettings')
      .evaluate()
      .setWidth(300)
      .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Confluence Settings');
}

function saveJIRAQuery(jql) {
  const querySheet = SpreadsheetApp.getActiveSheet();
  const queryMetadata = getSheetMetadataPropertyByKey(querySheet, 'jiraQuery');
  if (queryMetadata !== undefined) {
    queryMetadata.setValue(jql);
  } else {
    querySheet.addDeveloperMetadata('jiraQuery', jql);
  }
  addTriggersIfNotInstalled();
}

function loadAllPrevious_() {
  loadPrevious_(new Date(2020, 1, 26));
  loadPrevious_(new Date(2020, 1, 27));
  loadPrevious_(new Date(2020, 1, 28));
}

function updateConfluenceChart_(chart, confluencePageId, confluenceAttachmentId, confluenceAttachmentType) {
  confluenceAttachmentType = confluenceAttachmentType || MimeType.PNG;
  let filePrefix = 'chart';
  let chartBlob = Utilities.newBlob(chart.getAs(MimeType.PNG).getBytes(), confluenceAttachmentType, filePrefix + chart.getChartId() + '.png');
  return confluenceApiMultipartPost_('/api/content/' + confluencePageId + '/child/attachment/' + confluenceAttachmentId + '/data',
                                                   { minorEdit: 'true', comment: 'Updated from Google Sheets', file: chartBlob });
}

function addConfluenceChart_(chart, confluencePageId, confluenceAttachmentType) {
  confluenceAttachmentType = confluenceAttachmentType || MimeType.PNG;
  let filePrefix = 'chart';
  let chartBlob = Utilities.newBlob(chart.getAs(MimeType.PNG).getBytes(), confluenceAttachmentType, filePrefix + chart.getChartId() + '.png');
  return confluenceApiMultipartPost_('/api/content/' + confluencePageId + '/child/attachment',
                                                   { minorEdit: 'true', comment: 'Added from Google Sheets', file: chartBlob });
}

function saveChartToDrive_(chart, attachmentType) {
  attachmentType = attachmentType || MimeType.PNG;
  let filePrefix = SpreadsheetApp.getActiveSpreadsheet().getName() + ' ';
  let chartBlob = Utilities.newBlob(chart.getAs(MimeType.PNG).getBytes(), attachmentType, filePrefix + chart.getChartId() + '.png');
  DriveApp.createFile(chartBlob);
}

function getChartData_() {
  let chartData = [];
  let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data: Charts');
  if (dataSheet === null) {
    throw 'Charts data sheet was not found';
  }
  let chartRows = getTableData_(dataSheet);
  for (let i=0; i<chartRows.length; i++) {
    let sheetName = chartRows[i]['Sheet Name'], chartIndex = chartRows[i]['Chart Index'],
        confluencePageId = chartRows[i]['Confluence PageID'], confluenceAttachmentId = chartRows[i]['Confluence AttachmentID'],
        confluenceAttachmentType = chartRows[i]['File Type'];
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet === null) {
      throw 'Sheet ' + sheetName + ' was not found';
    }
    let pageCharts = sheet.getCharts();
    if (typeof chartIndex === 'number' && chartIndex >= 0 && chartIndex < pageCharts.length) {
      chartRows[i]['chart'] = pageCharts[chartIndex];
      chartData.push(chartRows[i]);
    } else {
      throw 'Chart with index ' + chartIndex + ' on sheet ' + sheetName + ' was not found';
    }
  }
  return chartData;
}

function setChartData_(rowIndex, columnValues) {
  const chartsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data: Charts'),
    sheetColumns = getTableHeaders_(chartsSheet);
  for (const key in columnValues) {
    if (columnValues.hasOwnProperty(key)) {
      const columnIndex = sheetColumns.indexOf(key);
      if (columnIndex > -1) {
        const columnValue = columnValues[key];
        chartsSheet.getRange(rowIndex + 1, columnIndex + 1).setValue(columnValue);
      }
    }
  }
}

function publishDriveCharts() {
  let chartData = getChartData_();
  chartData.forEach(function(dataItem) {
    saveChartToDrive_(dataItem['chart'], dataItem['File Type']);
  });
}

function updateConfluenceCharts() {
  let chartData = getChartData_();
  chartData.forEach(function(dataItem, chartIndex) {
    let pageId = dataItem['Confluence PageID'], fileType = dataItem['File Type'];
    if (typeof pageId === 'number') {
      let attachmentId = dataItem['Confluence AttachmentID'];
      if (typeof attachmentId === 'number') {
        updateConfluenceChart_(dataItem['chart'], pageId,
          attachmentId, fileType);
      } else {
        let attachResponse = addConfluenceChart_(dataItem['chart'], pageId, fileType);
        if (attachResponse['results'] && attachResponse['results'].length == 1) {
          setChartData_(chartIndex + 1, {'Confluence AttachmentID': attachResponse['results'][0]['id']});
        }
      }
    } else {
      throw 'Confluence PageId must be numeric';
    }
  });
}

function loadPrevious_(previousDate) {
  let todayStr = Utilities.formatDate(previousDate, "GMT", "yyyy-MM-dd");
  let newSheet = SpreadsheetApp.getActive().getSheetByName('Worksheet ' + todayStr);
  let destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All Issues');
  let nextColumn = findGivenDateColumn_(destSheet, previousDate);
  if (nextColumn === -1) {
    nextColumn = destSheet.getMaxColumns() + 1;
    destSheet.getRange(1, nextColumn).setValue(todayStr);
  }
  populateAllItems_(newSheet, destSheet, nextColumn);
}

function findGivenDateColumn_(sheet, now) {
  for (let column=1; column<=sheet.getMaxColumns(); column++) {
    let columnValue = sheet.getRange(1, column).getValues()[0][0];
    if (columnValue instanceof Date && columnValue.getFullYear() === now.getFullYear() &&
        columnValue.getMonth() === now.getMonth() && columnValue.getDate() === now.getDate()) {
          return column;
    }
  }
  return -1;
}

function findLastDateColumn_(sheet) {
  let lastColumnWithDateValue = 0;
  for (let column=1; column<=sheet.getMaxColumns(); column++) {
    let columnRange = sheet.getRange(1, column), columnValue = columnRange.getValue(), columnFormula = columnRange.getFormula();
    if (columnValue instanceof Date && columnFormula === '') {
      lastColumnWithDateValue = column;
    }
  }
  return lastColumnWithDateValue;
}

function addChartDateColumns() {
  let chartData = getChartData_();
  chartData.forEach(function(dataItem) {
    let chartSheet = SpreadsheetApp.getActive().getSheetByName(dataItem['Sheet Name']);
    addTodayColumn_(chartSheet);
  });
}

function addTodayColumn_(sheet) {
  let now = new Date();
  let todayDateColumn = findGivenDateColumn_(sheet, now);
  if (todayDateColumn < 1) {
    let lastColumnWithDateValue = findLastDateColumn_(sheet);
    if (lastColumnWithDateValue > 0) {
      let newColumnPos = lastColumnWithDateValue + 1, sheetMaxRows = sheet.getMaxRows();
      sheet.insertColumns(newColumnPos, 1);
      // Set date heading
      sheet.getRange(1, newColumnPos).setValue(dateAsString_(now));
      // Copy data in the rows
      if (sheetMaxRows > 1) {
        let sourceRange = sheet.getRange(2, lastColumnWithDateValue, sheetMaxRows - 1),
            destination = sheet.getRange(2, lastColumnWithDateValue, sheetMaxRows - 1, 2);
        sourceRange.autoFill(destination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
      }
    }
  }
}

function getJIRAIssues_(jql, newSheet) {
  if (jql === null || jql === '') {
    throw 'Sheet property jiraQuery was not found or empty';
  }
  let startAt = 0;
  let responseData = null;
  let newRows = [];
  do {
    responseData = jiraApiGet_('/api/2/search?jql=' + encodeURIComponent(jql) + '&startAt=' + startAt);
    for (let count=0; count<responseData.issues.length; count++) {
      let issue = responseData.issues[count];
      newRows.push([issue.fields.issuetype.name, issue.key, issue.fields.summary, issue.fields.customfield_11423, issue.fields.status.name,
                    issue.fields.assignee ? issue.fields.assignee.key : '', new Date(issue.fields.created), '']);
    }
    startAt += responseData.issues.length;
  } while(responseData.issues.length > 0)
  newSheet.getRange(1, 1, 1, 8).setValues([['Issue Type', 'Issue Key', 'Summary', 'Epic Key', 'Status', 'Assignee', 'Created', '']]);
  newSheet.getRange(2, 1, newRows.length, 8).setValues(newRows);
}

function jiraApiGet_(path, url=null, username=null, password=null) {
  username = username || getJIRAUsername(), password = password || getJIRAPassword();
  const restUrl = url = (url || getJIRAUrl()) + '/rest' + path;
  const response = UrlFetchApp.fetch(restUrl, {'muteHttpExceptions': false, 'headers': {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)}});
  return JSON.parse(response.getContentText());
}

function multipartContent_(fieldValues, boundary) {
  var payload = [];
  for (let field in fieldValues) {
    let fieldHeaders = "Content-Disposition: form-data; name=\"" + field + "\"";
    let fieldValue;
    if (typeof fieldValues[field].getBytes === 'function') {
      fieldHeaders = fieldHeaders + "; filename=\"" + fieldValues[field].getName() + "\"\r\n" + "Content-Type: " + fieldValues[field].getContentType();
      fieldValue = fieldValues[field];
    } else {
      fieldValue = Utilities.newBlob(fieldValues[field]);
    }
    payload = payload.concat(Utilities.newBlob("--" + boundary + "\r\n" + fieldHeaders + "\r\n\r\n").getBytes())
      .concat(fieldValue.getBytes())
      .concat(Utilities.newBlob("\r\n").getBytes());
  }
  payload = payload.concat(Utilities.newBlob("--" + boundary + "--\r\n").getBytes());
  return payload;
}

function confluenceApiMultipartPost_(path, data) {
  const boundary = '------------------------' + Utilities.getUuid().replace('-', '').substr(0, 16);
  const requestContentType = 'multipart/form-data; boundary=' + boundary;
  return confluenceApiRequest_(path, multipartContent_(data, boundary), requestContentType);
}

function confluenceApiRequest_(path, data=null, contentType=null, url=null, username=null, password=null) {
  username = username || getConfluenceUsername(), password = password || getConfluencePassword();
  const restUrl = url = (url || getConfluenceUrl()) + '/rest' + path;
  const requestHeaders = {
    'Authorization': 'Basic ' + Utilities.base64Encode(username + ':' + password),
  };
  const fetchParameters = {
    muteHttpExceptions: false,
    headers: requestHeaders,
  };
  if (data !== null) {
    fetchParameters['method'] = 'post';
    fetchParameters['payload'] = data;
    if (contentType !== null) {
      fetchParameters['contentType'] = contentType;
    }
    requestHeaders['X-Atlassian-Token'] = 'nocheck';
  }
  let response = UrlFetchApp.fetch(restUrl, fetchParameters);
  return JSON.parse(response.getContentText());
}

function getJIRAHistoryForIssue_(issueKey) {
  let responseData = jiraApiGet_('/api/2/issue/' + issueKey + '?expand=changelog');
  let statusValues = [];
  let oldStatusValues = [];
  let changelog = responseData['changelog']['histories'], totalResults = responseData['changelog']['total'], items = [];
  for (let i=0; i<totalResults; i++) {
    items = changelog[i]['items'];
    for (let j=0; j<items.length; j++) {
      if (items[j]['field'] === 'status') {
        oldStatusValues.push([changelog[i]['created'], items[j]['fromString']]);
        statusValues.push([changelog[i]['created'], items[j]['toString']]);
      }
    }
  }
  
  // Infer original status from first transition or (if no transitions) use current value
  let firstTransitionOldValue = oldStatusValues.length > 0 ? oldStatusValues[0] : null;
  let createdDate = responseData['fields']['created'];
  let currentStatusValue = responseData['fields']['status']['name'];
  if (firstTransitionOldValue) {
    statusValues = [[createdDate, firstTransitionOldValue[1]]].concat(statusValues);
  } else {
    statusValues = [[createdDate, currentStatusValue]];
  }
  
  return statusValues;
}

function getIssueStatusForDates_(issueKey, dateValues) {
  let transitions = getJIRAHistoryForIssue_(issueKey);
  console.log(transitions);
  
  let currentStatusIndex = -1;
  return dateValues.map(function(currentDateTime) {
    let currentDateStr = currentDateTime.getFullYear() + '-' + Utilities.formatString('%02d', currentDateTime.getMonth() + 1) + '-' + Utilities.formatString('%02d', currentDateTime.getDate());
    while(currentStatusIndex < transitions.length-1 && transitions[currentStatusIndex + 1][0].substr(0, 10) < currentDateStr) {
      currentStatusIndex ++;
    }
    return currentStatusIndex >= 0 ? transitions[currentStatusIndex][1] : '';
  });
}

function fillIssueHistory_(itemsSheet, rowNumber) {
  let dateValues = itemsSheet.getRange(1, 5, 1, itemsSheet.getLastColumn()-4).getValues()[0];
  let issueKey = itemsSheet.getRange(rowNumber + 1, 2, 1, 1).getValue();
  let statusValues = getIssueStatusForDates_(issueKey, dateValues);
  itemsSheet.getRange(rowNumber + 1, 5, 1, dateValues.length).setValues([statusValues]);
}

function fillSheetIssueHistories(itemsSheet=null) {
  itemsSheet = itemsSheet || SpreadsheetApp.getActiveSheet();
  let startRow = 1;
  let lastRow = itemsSheet.getLastRow();
  for (let row=startRow; row<=lastRow-1; row++) {
    fillIssueHistory_(itemsSheet, row);
  }
}

function populateAllItems_(sourceSheet, destSheet, insertColumnIndex) {
  addMissingItems_(sourceSheet, destSheet);
  let sourceData = getTableData_(sourceSheet);
  let destData = getTableData_(destSheet);
  let tableColumnValues = getTableColumnData_(sourceData, destData, 'Issue Key', 'Status');
  destSheet.getRange(2, insertColumnIndex, tableColumnValues.length, 1).setValues(tableColumnValues);
}

function addMissingItems_(sourceSheet, dstSheet) {
  var translations = {'Issue Key': 'Issue Key'}; // dst => src
  var dstCols = getTableHeaders_(dstSheet);
  var missingItems = getMissingIssues_(getTableData_(sourceSheet), getTableData_(dstSheet));
  var rowsToAdd = missingItems.map(itemToAdd => { return dstCols.map((colName) => itemToAdd[translations[colName] || colName]) });
  if (rowsToAdd.length) {
    var dstData = getTableData_(dstSheet);
    dstSheet.getRange(dstData.length + 2, 1, rowsToAdd.length, dstCols.length).setValues(rowsToAdd);
  }
}

function getMissingIssues_(rowItems, existingItems) {
  return rowItems.filter((row) => !getMatchingRow_(row, existingItems, 'Issue Key'));
}

function getTableHeaders_(currentSheet) {
  let allColumns = currentSheet.getRange(1, 1, 1, currentSheet.getMaxColumns()).getValues()[0].map((item => typeof item === 'string' ? item : ''));
  let firstBlank = allColumns.indexOf('');
  if (firstBlank > -1) {
    return allColumns.slice(0, firstBlank);
  } else {
    return allColumns;
  }
}

function getTableData_(currentSheet) {
  var keys = getTableHeaders_(currentSheet);
  var data = currentSheet.getRange(2, 1, currentSheet.getMaxRows() - 1, keys.length).getValues();
  return data.filter((row) => row.length > 0 && row[0]).map(function(row, rowIndex) {
    var rowValues = {};
    keys.forEach((key, colIndex) => rowValues[key] = row[colIndex]);
    return rowValues;
  });
}

function getTableColumnData_(sourceTableData, destSheetData, columnName, outputColumnName) {
  let columnValues = destSheetData.map((destRow) => { var matchRow = getMatchingRow_(destRow, sourceTableData, columnName); return matchRow ? matchRow[outputColumnName] : '' });
  return columnValues.map((item) => [item]);
}

function getMatchingRow_(needle, haystack, propName) {
  return haystack.find((haystackRow) => {
    return haystackRow[propName] !== '' && haystackRow[propName] === needle[propName]
  });
}

function getJIRAUsername() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('jiraUsername');
}

function getJIRAPassword() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('jiraPassword');
}

function testJIRACredentials(url, username, password) {
  return jiraApiGet_('/api/2/myself', url, username, password);
}

function saveJIRACredentials(url, username, password) {
  const properties = PropertiesService.getUserProperties(),
    scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('jiraUrl', url);
  properties.setProperty('jiraUsername', username);
  properties.setProperty('jiraPassword', password);
}

function getConfluenceUrl() {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty('confluenceUrl');
}

function getConfluenceUsername() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('confluenceUsername');
}

function getConfluencePassword() {
  const properties = PropertiesService.getUserProperties();
  return properties.getProperty('confluencePassword');
}

function testConfluenceCredentials(url, username, password) {
  return confluenceApiRequest_('/api/user/current', null, null, url=url, username=username, password=password);
}

function saveConfluenceCredentials(url, username, password) {
  const properties = PropertiesService.getUserProperties(),
    scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('confluenceUrl', url);
  properties.setProperty('confluenceUsername', username);
  properties.setProperty('confluencePassword', password);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
