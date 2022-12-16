const headers = {
    "content-type": "application/json",
    "Accept": "application/json",
    "Authorization": "Bearer " + ''
  };

const jiraParams = {
    "method": "GET",
    "headers": headers
  };

const jiraQueryBase = "https://issues.redhat.com/rest/api/2/search";
const driveId = "";
const chartDrawingSheet = '';

const main = () => {
  const templateDoc = DocumentApp.getActiveDocument();  
  const newDocId = copyTemplate(templateDoc.getId());
  const newDoc = DocumentApp.openById(newDocId);

  const placeholders = findPlaceholders(newDoc.getBody().getText());

  Logger.log('Found ' + placeholders.length + ' jql plceholders, going to execute them.');
  for (const placeholder of placeholders) {

    Logger.log('Processing placeholder');
    if (placeholder.type == 'CFD') {
      const loadedIssues = loadJiraQuery(placeholder.queryString, "expand=changelog&");
      const cfdData = translateToCfd(loadedIssues);
      const spreadsheetFriendlyMap = rawMapToSpreadsheetMap(cfdData);
      writeJiraDataToCfd(spreadsheetFriendlyMap);
      Logger.log('Processed, copying to the output doc');

      copyChartFromSpreadsheet('Input', newDoc.getBody(), placeholder.raw);
      Logger.log('Copied');
    } else if (placeholder.type == 'PIE') {
      const loadedIssues = loadJiraQuery(placeholder.queryString, '');
      const pieData = translateToPieChart(loadedIssues, placeholder.params);
      writeJiraDataToPie(pieData);
      Logger.log('Processed, copying to the output doc');
      copyChartFromSpreadsheet('PieChart', newDoc.getBody(), placeholder.raw);
      Logger.log('Copied');
    }
  }

  Logger.log('Done, report is at: https://docs.google.com/document/d/' + newDocId);
};

const pieMain = (query) => {
  const loadedIssues = loadJiraQuery(query, '');
  const pieData = translateToPieChart(loadedIssues);
  writeJiraDataToPie(pieData);
};

const copyTemplate = (templateId) => {
  const getDate = () => {
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();

    return mm + '/' + dd + '/' + yyyy;
  }

  Logger.log('Copying template to output report.');
  const newDoc = DocumentApp.create('[LATEST] Ecosystem UI Weekly Report - ' + getDate());
  const newDocId = newDoc.getId();
  const newDocFile = DriveApp.getFileById(newDocId);
  mergeDocs([newDocId, templateId], false);
  DriveApp.getFolderById(driveId).addFile(newDocFile);
  DriveApp.getRootFolder().removeFile(newDocFile);
  Logger.log('Template copied copied');

  return newDocId;
};

const findPlaceholders = (bodyText) => {
    const regexpNames =  /###inserFromJQL: (.*); (.*); (.*)/mg;
    const res = [];
    for (const match of bodyText.matchAll(regexpNames)) {
      const cleanedQuery = match[3].replaceAll('‘', '\'').replaceAll('’', '\'').replaceAll('"', '\'').trim();
      const raw = '###inserFromJQL: ' + match[1];
      res.push({'raw': raw, 'type': match[1], 'params': match[2].trim(), queryString: cleanedQuery});
    }

    return res;
};

const writeJiraDataToCfd = (data) => {
  const sheet = SpreadsheetApp.openById(chartDrawingSheet).getSheetByName('Input');

  const res = [];
  for (let i in data) {
    let innerRes = [i];
    for (state of Object.values(outputStates)) {
      innerRes.push(data[i][state]);
    }
    res.push(innerRes);
  }
  
  sheet.getRange(2, 1, sheet.getLastRow(), 5).clearContent();
  if (res.length != 0) {
    sheet.getRange(2, 1, res.length, res[0].length).setValues(res);
  }
};

// this function is taken from https://gist.github.com/tanaikech/f84831455dea5c394e48caaee0058b26
 const replaceTextToImage = (body, searchText, image, width) => {
    const next = body.findText(searchText);
    if (!next) {
      return;
    }

    const r = next.getElement();
    r.asText().setText("");
    const img = r.getParent().asParagraph().insertInlineImage(0, image);
    if (width && typeof width == "number") {
      const w = img.getWidth();
      const h = img.getHeight();
      img.setWidth(width);
      img.setHeight(width * h / w);
    }
    return next;
  };

const copyChartFromSpreadsheet = (inputSheet, body, textToReplace) => {
  const spreadsheet = SpreadsheetApp.openById(chartDrawingSheet);
  const charts = spreadsheet.getSheetByName(inputSheet).getCharts();  

  for (let i in charts) {
    replaceTextToImage(body, textToReplace, charts[i]);
  }
};

const outputStates = {
  DONE: 'Done',
  ON_QE: 'On QE',
  IN_PROGRESS: 'In Progress',
  BACKLOG: 'Backlog'
};

const inputToOutputStates = {
  'To Do': outputStates.BACKLOG,
  'Planning': outputStates.BACKLOG,
  'New': outputStates.BACKLOG,
  'ASSIGNED': outputStates.BACKLOG,
  
  'In Progress': outputStates.IN_PROGRESS,
  'Dev Complete': outputStates.IN_PROGRESS,
  'POST': outputStates.IN_PROGRESS,
  'Code Review': outputStates.IN_PROGRESS,
  
  'Review': outputStates.ON_QE,
  'ON_QA': outputStates.ON_QE,
  'Feature Complete': outputStates.ON_QE,
  'QE Review': outputStates.ON_QE,
  'MODIFIED': outputStates.ON_QE,
  
  'Closed': outputStates.DONE,
  'Verified': outputStates.DONE,
  'Release Pending': outputStates.DONE,
  'Done': outputStates.DONE,
  'Obsolete': outputStates.DONE,
  "Won't Fix / Obsolete": outputStates.DONE,
};

const addToResult = (res, weekNumber, toState, issueKey) => {
  if (!(weekNumber in res)) {
    res[weekNumber] = {[outputStates.BACKLOG]: [], [outputStates.IN_PROGRESS]: [], [outputStates.ON_QE]: [], [outputStates.DONE]: []};
  }

  for (state of Object.values(outputStates)) {
    const issuesInState = res[weekNumber][state];
    const issueIndex = issuesInState.indexOf(issueKey);
    if (issueIndex > -1) {
      issuesInState.splice(issueIndex, 1);
    }
  }

  const stateArr = res[weekNumber][toState];
  if (stateArr.indexOf(issueKey) === -1) {
    stateArr.push(issueKey);
  }
};

const rawMapToSpreadsheetMap = (raw) => {
  const res = [];

  const emptyRow = () => {
    return {[outputStates.BACKLOG]: 0, [outputStates.IN_PROGRESS]: 0, [outputStates.ON_QE]: 0, [outputStates.DONE]: 0};
  };

  const containsInAnyState = (issue, week) => {
    for (containsInStateVal of Object.values(outputStates)) {
        // this condition does not work!
        if (week[containsInStateVal].includes(issue)) {
          return true;
        }
      }

    return false;
  };
  
  const sortedWeeks = Object.keys(raw);

  const min = Math.min(...sortedWeeks);
  const max = Math.max(...sortedWeeks);
  let prevWeek = {[outputStates.BACKLOG]: [], [outputStates.IN_PROGRESS]: [], [outputStates.ON_QE]: [], [outputStates.DONE]: []};

  // need to iterate this way since not all the weeks need to actually exist, so the gaps need to be filled
  // for example min can be 2, max 10, but in between no other weeks have any changes. This loop goes over all weeks though filling the details in
  for (let week = min; week <= max; week ++) {
    const newRow = emptyRow();
    if (week in raw) {
      let rawWeek = raw[week];
      for (val of Object.values(outputStates)) {
        for (let issue of prevWeek[val]) {
          if (!containsInAnyState(issue, rawWeek)) {
            rawWeek[val].push(issue);
          }
        }
      }
      
      prevWeek = rawWeek;
    }

    for (outputVal of Object.values(outputStates)) {
        newRow[outputVal] = prevWeek[outputVal].length;
    }
    
    res.push(newRow);
  }

  return res;
};

const loadJiraQuery = (query, expand) => {
  Logger.log('Trying to load issues for query: ' + query);
  
  let start = 0;
  const limit = 500;
  let total = 0;
  const allIssues = [];
  do {
    const fullQueryUrl = jiraQueryBase + "?" + expand + "maxResults=" + limit + "&startAt=" + start + "&jql=" + query;
    Logger.log('Launching paginated query: ' + fullQueryUrl);

    const response = UrlFetchApp.fetch(fullQueryUrl, jiraParams);
    const content = JSON.parse(response.getContentText());
    
    start = start + limit;
    total = content['total'];
    allIssues.push(...content['issues']);

  } while (total >= start);

  Logger.log('Number of issues loaded to process: ' + allIssues.length);
  
  return allIssues;
};

const translateToCfd = (allIssues) => {
  // returns the number of weeks since the unix epoch
  const toWeek = (str) => Math.ceil(new Date(str).getTime() / 1000 / 604800);

  const res = {};
  for (const issue of allIssues) {
    const createdWeek = toWeek(issue['fields']['created']);
    // this can be counted twice if it changed within the same week
    addToResult(res, createdWeek, outputStates.BACKLOG, issue['key']);
    // number of weeks since the beginning of epoch
    const submits = issue['changelog']['histories'];

    for (const submit of submits) {
      const submitTimestamp = submit['created'];
      
      // every time someone hits "save" on the issue, one or more items can be changed
      for (const itemChanged of submit['items']) {
        const fieldCanged = itemChanged['field'];
        if (fieldCanged == 'status') {
          const weekOfChange = toWeek(submitTimestamp);
          addToResult(res, weekOfChange, inputToOutputStates[itemChanged['toString']], issue['key']);
        }
      }
    }
  }
  
  return res;
};

const translateToPieChart = (allIssues, path) => {
  const res = {};
  for (const issue of allIssues) {
    const value = parseValueFromIssue(issue, path);
    if (!(value in res)) {
      res[value] = 0;
    }

    res[value] ++;
  }

  return res;
};

const writeJiraDataToPie = (data) => {
  const sheet = SpreadsheetApp.openById(chartDrawingSheet).getSheetByName('PieChart');

  const res = [];
  for (let key in data) {
    res.push([key, data[key]]);
  }
  
  sheet.getRange(2, 1, sheet.getLastRow(), 2).clearContent();
  if (res.length != 0) {
    sheet.getRange(2, 1, res.length, res[0].length).setValues(res);
  }
};

const parseValueFromIssue = (issue, path) => {
  let res = issue;
  for (key of path.split(".")) {
    // todo - if the last one is a list, squash the values to one
    res = res[key]
  }
  
  return res;
};

// shamelessly copied this function from: https://stackoverflow.com/questions/17575863/how-do-i-format-text-i-am-copying-from-google-document-in-google-app-script
function mergeDocs(docIDs,pagebreak) {
  var baseDoc = DocumentApp.openById(docIDs[0]);
  var body = baseDoc.getBody();

  for( var i = 1; i < docIDs.length; ++i ) {
    if (pagebreak) body.appendPageBreak();
    var otherBody = DocumentApp.openById(docIDs[i]).getBody(); 
    var totalElements = otherBody.getNumChildren();
    var latestElement;
    for( var j = 0; j < totalElements; ++j ) {
      var element = otherBody.getChild(j).copy();
      var attributes = otherBody.getChild(j).getAttributes();
      var type = element.getType(); 
      if (type == DocumentApp.ElementType.PARAGRAPH) {
        if (element.asParagraph().getNumChildren() != 0 && element.asParagraph().getChild(0).getType() == DocumentApp.ElementType.INLINE_IMAGE) {
          var pictattr = element.asParagraph().getChild(0).asInlineImage().getAttributes();
          var blob = element.asParagraph().getChild(0).asInlineImage().getBlob();
          // Image attributes, e.g. size, do not survive the copy, and need to be applied separately
          latestElement = body.appendImage(blob);
          latestElement.setAttributes(clean(pictattr));
        }
        else latestElement = body.appendParagraph(element);
      }
      else if( type == DocumentApp.ElementType.TABLE )
        latestElement = body.appendTable(element);
      else if( type == DocumentApp.ElementType.LIST_ITEM )
        latestElement = body.appendListItem(element);
      else
        throw new Error("Unsupported element type: "+type);
      latestElement.setAttributes(clean(attributes));
    }
  }
}

/**
 * Remove null attributes in style object, obtained by call to
 * .getAttributes().
 * https://code.google.com/p/google-apps-script-issues/issues/detail?id=2899
 */
function clean(style) {
  for (var attr in style) {
    if (style[attr] == null) delete style[attr];
  }
  return style;
}

// -------------------------------------------------------------------------
// I know, this should be a new file
// -------------------------------------------------------------------------

function assert(condition, message) {
    if (!condition) {
        throw message || "Assertion failed";
    }
}

const testAll = () => {
  testRawMapToSpreadsheetMap();
  testFindPlaceholders();
  testAddToResult();
  testParseValueFromIssue();

  Logger.log('All tests passed!');
};

const testParseValueFromIssue = () => {
  assert(parseValueFromIssue({'someKey': 'someVal'}, 'someKey') == 'someVal');
  assert(parseValueFromIssue({'someKey': {'nestedKey': 'nestedVal'}}, 'someKey.nestedKey') == 'nestedVal');
  assert(parseValueFromIssue({'someKey': {'nestedKey': {'third': 'thirdVal'}}}, 'someKey.nestedKey.third') == 'thirdVal');
};

const testRawMapToSpreadsheetMap = () => {
  const input = {
    5 : {
      "In Progress": [],
      "On QE": [],
      "Done": [],
      "Backlog": ['t1', 't2']
      },
    7 : {
      "In Progress": ['t1'],
      "On QE": [],
      "Done": [],
      "Backlog": []
      },
    10 : {
      "In Progress": [],
      "On QE": ['t2'],
      "Done": [],
      "Backlog": []
      },

    11 : {
      "In Progress": [],
      "On QE": [],
      "Done": ['t4'],
      "Backlog": []
      }
  }

  const res = rawMapToSpreadsheetMap(input);
  assert(res.length == 7);

  assert(res[3]['In Progress'] == 1);
  assert(res[3]['On QE'] == 0);
  assert(res[3]['Done'] == 0);
  assert(res[3]['Backlog'] == 1);

  assert(res[6]['In Progress'] == 1);
  assert(res[6]['On QE'] == 1);
  assert(res[6]['Done'] == 1);
  assert(res[6]['Backlog'] == 0);

  writeJiraDataToCfd(res);
};

const testAddToResult = () => {
  const res = {};
  
  // add two in one week
  addToResult(res, 5, outputStates.BACKLOG, 'key-1');
  addToResult(res, 5, outputStates.BACKLOG, 'key-2');
  assert(res[5][outputStates.BACKLOG].length == 2);
  assert(res[5][outputStates.DONE].length == 0);
  assert(res[5][outputStates.ON_QE].length == 0);
  assert(res[5][outputStates.IN_PROGRESS].length == 0);

  // add the same in the same week
  addToResult(res, 5, outputStates.IN_PROGRESS, 'key-2');
  assert(res[5][outputStates.BACKLOG].length == 1);
  assert(res[5][outputStates.DONE].length == 0);
  assert(res[5][outputStates.ON_QE].length == 0);
  assert(res[5][outputStates.IN_PROGRESS].length == 1);

  // add to a different week
  addToResult(res, 6, outputStates.DONE, 'key-2');
  assert(res[5][outputStates.BACKLOG].length == 1);
  assert(res[5][outputStates.DONE].length == 0);
  assert(res[5][outputStates.ON_QE].length == 0);
  assert(res[5][outputStates.IN_PROGRESS].length == 1);

  assert(res[6][outputStates.BACKLOG].length == 0);
  assert(res[6][outputStates.DONE].length == 1);
  assert(res[6][outputStates.ON_QE].length == 0);
  assert(res[6][outputStates.IN_PROGRESS].length == 0);
};

const testFindPlaceholders = () => {
    const bodyText = `
     
something els

gla

###inserFromJQL: CFD; ; component = 'MGMT UI' and fixVersion = RHACM-2.6
something

###inserFromJQL: PIE; fields.status.name ; component = 'Assisted-Installer UI'

bla after
`;

  const res = findPlaceholders(bodyText);
  assert(res[0].queryString == "component = 'MGMT UI' and fixVersion = RHACM-2.6");
  assert(res[0].type == "CFD");
  assert(res[0].params == "");

  assert(res[1].queryString == "component = 'Assisted-Installer UI'");
  assert(res[1].type == "PIE");
  assert(res[1].params == "fields.status.name");
  assert(res.length == 2);

};
