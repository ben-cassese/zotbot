const library_key = 'redacted';
const collection_key = 'redacted';
const api_key = 'redacted';

// Setup the triggers for the sheet
 var triggers = ScriptApp.getProjectTriggers();
 for (var i = 0; i < triggers.length; i++) {
   ScriptApp.deleteTrigger(triggers[i]);
 }

ScriptApp.newTrigger("menus")
   .forSpreadsheet(SpreadsheetApp.getActive())
   .onOpen()
   .create();

ScriptApp.newTrigger("link_dashboard_notes(e)")
   .forSpreadsheet(SpreadsheetApp.getActive())
   .onEdit()
   .create();


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Setup the menu
function menus() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ZotBot')
      .addItem('Pull from Zotero', 'pull')
      .addItem('Push changes to Zotero', 'push')
      .addItem('Setup Everything', 'setup_everything')
      .addToUi();
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Setup the whole thing
function setup_everything(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('ZotBot Setup', 'This will reset/setup the whole sheet. Are you sure you want to proceed?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('Paper Notes');
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('Zotero Import');
    SpreadsheetApp.getActiveSpreadsheet().insertSheet('Differences');
    pull();
    dashboard_setup();
  } else {
    return;
  }

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Pull Data from Zotero, setup the notes sheet
function pull(){
  fetch_Zot_Data();
  setup_notes_from_Zot();
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Compile then push changes to Zotero
function push(){
  var t = log_changes();
  var notes = t[0];
  for (let i = 0; i < notes[0].length; i++) {
    var item_key = notes[0][i][0];
    var content = notes[1][i][0];
    update_zot_item(item_key, content);
  }
  var priorities = t[1];
  for (let i = 0; i < priorities[0].length; i++) {
    var item_key = priorities[0][i][0];
    var content = '@@Priority: ' + priorities[1][i][0].toString() + '@@';
    update_zot_item(item_key, content);
  }
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Run all the time on edit
function link_dashboard_notes(e){
  if (e.range.getA1Notation() !== 'C9') {
    return;
  }
  var dash_sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  var notes_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paper Notes');

  var active_row = dash_sheet.getRange("E4").getValue();
  var notes = dash_sheet.getRange("C9").getValue();
  notes_sheet.getRange(active_row, 6).setValue(notes + ' ##ZotBot##');
}



/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////// HELPERS ///////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Create the initial notes for a Zotero item that will get pulled here
function create_ZotBot_notes(paper_key){
  var response = UrlFetchApp.fetch(url='https://api.zotero.org/items/new?itemType=note');
  const note_template = JSON.parse(response.getContentText());
  var zotbot_note = JSON.parse(JSON.stringify(note_template));
  zotbot_note['note'] = '##ZotBot##';
  var priority_note = JSON.parse(JSON.stringify(note_template));
  priority_note['note'] = '@@Priority: 999@@';


  var headers = {
  "Zotero-Write-Token": Utilities.getUuid().split('-').join(''),
  "Content-Type": "application/json"};

  var options = {
  'method' : 'post',
  'contentType': 'application/json',
  'payload' : JSON.stringify([zotbot_note, priority_note]),
  'headers' : headers};

  var response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items?key=' + api_key, options);
  var written = JSON.parse(response.getContentText());
  var note_key = written["success"][0];
  var priority_key = written["success"][1];

  options = {
  'method' : 'patch',
  'payload' : JSON.stringify({"parentItem" : paper_key}),
  'headers':{"If-Unmodified-Since-Version": response.getHeaders()["last-modified-version"]}};
  response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + note_key +'?key=' + api_key, options);
  response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + priority_key +'?key=' + api_key, options);
  
  return [note_key, priority_key];
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Pull data from Zotero
function fetch_Zot_Data() {
  SpreadsheetApp.getActive().toast('Pulling data from Zotero', "ZotBot Test");
  // Set the headers
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Zotero Import');
  var headers = [["Zotero Paper Key",	"Paper Title", "Date",	"Author",
  	"Note Zotero Key",	"Notes",	"Priority Zotero Key",	"Priority",	"URL",	"Abstract",	"Read", "Internal Ind"]];
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);

  // Get the first 100 papers
  var url = 'https://api.zotero.org/users/' + library_key +'/collections/' + collection_key + '/items/top?key=' + api_key + '&limit=100' ;
  var response = UrlFetchApp.fetch(url);
  var papers = JSON.parse(response.getContentText());

  // Get all of the other papers, if there are more than 100
  var start = 100;
  const num_items = response.getAllHeaders()['total-results'];
  while (num_items > start){
    response = UrlFetchApp.fetch(url + '&start=' + start.toString());
    papers.push(...JSON.parse(response.getContentText()));
    start += 100;
  }

  // Set the columns you can get from the parent item
  var internal_index = [];
  var paper_keys = [];
  var authors = [];
  var titles = [];
  var dates = [];
  var urls = [];
  var abstracts = [];
  for (let i = 0; i < papers.length; i++) {
    internal_index.push([i]);
    paper_keys.push([papers[i]['key']]);
    authors.push([papers[i]['meta']['creatorSummary']]);
    titles.push([papers[i]['data']['title']]);
    dates.push([papers[i]['data']['date']]);
    urls.push([papers[i]['data']['url']]);
    abstracts.push([papers[i]['data']['abstractNote']]);
    }

  sheet.getRange(2, 1, paper_keys.length, paper_keys[0].length).setValues(paper_keys);
  sheet.getRange(2, 2, paper_keys.length, paper_keys[0].length).setValues(titles);
  sheet.getRange(2, 3, paper_keys.length, paper_keys[0].length).setValues(dates);
  sheet.getRange(2, 4, paper_keys.length, paper_keys[0].length).setValues(authors);
  sheet.getRange(2, 9, paper_keys.length, paper_keys[0].length).setValues(urls);
  sheet.getRange(2, 10, paper_keys.length, paper_keys[0].length).setValues(abstracts);

  var note_keys = [];
  var notes = [];
  var read = [];
  var priority_keys = [];
  var priorities =[];

  for (let i = 0; i < papers.length; i++) {
    var paper_key = papers[i]['key'];

    var url = 'https://api.zotero.org/users/' + library_key + '/items/' + paper_key + '/children' + '?key=' + api_key ;
    response = UrlFetchApp.fetch(url);
    var children = JSON.parse(response.getContentText());
    var note_key = '';
    var priority_key = '';

    // Check if the custom notes have already been added
    for (let i = 0; i < children.length; i++) {
      if (children[i]['data']['itemType'] == 'note') {
        if (children[i]['data']['note'].includes('##ZotBot##')) {
          note_key = children[i]['key'];
          note_keys.push([note_key]);
          notes.push([children[i]['data']['note']]);

          // Check if it's read or unread
          if (children[i]['data']['note'].length > '##ZotBot##'.length){
            read.push([1]);
          } else{
            read.push([0]);
          }
        }else if(children[i]['data']['note'].includes('@@Priority')){
          priority_key = children[i]['key'];
          priority_keys.push([priority_key]);
          priorities.push([Number(children[i]['data']['note'].split('@@Priority: ')[1].split('@@')[0])]);
        }
      }
    }

    // If the file has been added to Zotero since the last time this ran, add the notes
    if (note_key == ''){
      var k = create_ZotBot_notes(paper_key);
      note_key = k[0];
      note_keys.push([note_key]);
      notes.push(['##ZotBot##']);
      priority_key = k[1];
      priority_keys.push([priority_key]);
      priorities.push([999]);
      read.push([0]);
    }
  }
  sheet.getRange(2, 5, note_keys.length, note_keys[0].length).setValues(note_keys);
  sheet.getRange(2, 6, note_keys.length, note_keys[0].length).setValues(notes);
  sheet.getRange(2, 7, note_keys.length, note_keys[0].length).setValues(priority_keys);
  sheet.getRange(2, 8, note_keys.length, note_keys[0].length).setValues(priorities);
  sheet.getRange(2, 11, note_keys.length, note_keys[0].length).setValues(read);
  sheet.getRange(2, 12, note_keys.length, note_keys[0].length).setValues(internal_index);

  SpreadsheetApp.getActive().toast('Data pull complete', "ZotBot Test");

}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Reset notes sheet based on Zotero data
function setup_notes_from_Zot(){
  // Setup
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paper Notes');
  sheet.getRange("A:Z").clearContent();
  sheet.clearFormats();

  // Copy the data
  var zotdata = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Zotero Import').getDataRange().getValues();
  sheet.getRange(1,1, zotdata.length, zotdata[0].length).setValues(zotdata);
  var r = sheet.getRange(2,1, zotdata.length-1, zotdata[0].length);
  r.sort([{column: 11, ascending: true}, {column: 8, ascending: true}]);

  // Clean it up
  sheet.setFrozenRows(1);
  sheet.hideColumns(1);
  sheet.hideColumns(5);
  sheet.hideColumns(7);
  sheet.setColumnWidths(2, 1, 300);
  sheet.setColumnWidths(3, 1, 75);
  sheet.setColumnWidths(6, 1, 500);
  sheet.setColumnWidths(9, 1, 75);
  sheet.setColumnWidths(10, 1, 75);
  sheet.setColumnWidths(11, 1, 50);
  sheet.setRowHeight(1, 50);


  sheet.getRange(2, 2, zotdata.length).setWrap(true);
  sheet.getRange(2, 4, zotdata.length).setWrap(true);
  sheet.getRange(2, 6, zotdata.length).setWrap(true);
  sheet.getRange(2, 9, zotdata.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, 10, zotdata.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);


  sheet.getRange("A:Z").setFontFamily('Cambria');

  // Set the colors
  sheet.getRange(1,1, 1, zotdata[0].length).setBackground('#bf9000');
  sheet.getRange(1,1, 1, zotdata[0].length).setFontSize(14);
  sheet.getRange(1,1, 1, zotdata[0].length).setFontWeight("bold");
  for (let i = 2; i < zotdata.length+1; i++) {
    if (i%2 == 0){
      var r = sheet.getRange(i,1,1,zotdata[0].length);
      r.setBackground('#434343');
      r.setFontColor('white');
    } else{
      sheet.getRange(i,1,1,zotdata[0].length).setBackground("#cccccc");
    }

  }  
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Compile changes between notes sheet and Zotero sheet
function log_changes(){
  var diff_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Differences');
  var note_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paper Notes');
  var zot_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Zotero Import');
  diff_sheet.getRange("A:Z").clearContent();

  // Notes
  note_sheet.getRange("L:L").copyTo(diff_sheet.getRange("A:A"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  note_sheet.getRange("F:F").copyTo(diff_sheet.getRange("B:B"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  var r = diff_sheet.getRange(2,1,diff_sheet.getLastRow(),2);
  r.sort({column: 1, ascending: true});
  zot_sheet.getRange("F:F").copyTo(diff_sheet.getRange("C:C"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  var note_item_keys = [];
  var old_notes = [];
  var new_notes = [];
  for (let i = 2; i <= diff_sheet.getLastRow(); i++) {
    var a = diff_sheet.getRange(i,2).getValue();
    var b = diff_sheet.getRange(i,3).getValue();
    if (a!=b){
      note_item_keys.push([zot_sheet.getRange(i,5).getValue()]);
      old_notes.push([b]);
      new_notes.push([a]);
    }
  }
  diff_sheet.getRange("A:Z").clearContent();
  diff_sheet.clearFormats();

  // Priorities
  note_sheet.getRange("L:L").copyTo(diff_sheet.getRange("A:A"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  note_sheet.getRange("H:H").copyTo(diff_sheet.getRange("B:B"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  var r = diff_sheet.getRange(2,1,diff_sheet.getLastRow(),2);
  r.sort({column: 1, ascending: true});
  zot_sheet.getRange("H:H").copyTo(diff_sheet.getRange("C:C"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);
  var priority_item_keys = [];
  var old_priorities = [];
  var new_priorities = [];
  for (let i = 2; i <= diff_sheet.getLastRow(); i++) {
    var a = diff_sheet.getRange(i,2).getValue();
    var b = diff_sheet.getRange(i,3).getValue();
    if (a!=b){
      priority_item_keys.push([zot_sheet.getRange(i,7).getValue()]);
      old_priorities.push([b]);
      new_priorities.push([a]);
    }
  }
  diff_sheet.getRange("A:Z").clearContent();
  diff_sheet.clearFormats();

  // Write them out
  diff_sheet.getRange(1, 1, 1, 3).setValues([["Note Key", "Old Note", "New Note"]]);
  diff_sheet.setColumnWidths(2, 1, 300);
  diff_sheet.setColumnWidths(3, 1, 500);

  if (note_item_keys.length > 0){
    diff_sheet.getRange(2, 2, note_item_keys.length).setWrap(true);
    diff_sheet.getRange(2, 3, note_item_keys.length).setWrap(true);
    diff_sheet.getRange(2, 1, note_item_keys.length, note_item_keys[0].length).setValues(note_item_keys);
    diff_sheet.getRange(2, 2, note_item_keys.length, note_item_keys[0].length).setValues(old_notes);
    diff_sheet.getRange(2, 3, note_item_keys.length, note_item_keys[0].length).setValues(new_notes);
  }

  diff_sheet.getRange(1, 5, 1, 3).setValues([["Priority Key", "Old Priority", "New Priority"]]);
  diff_sheet.setColumnWidths(5, 1, 110);
  diff_sheet.setColumnWidths(6, 1, 110);
  diff_sheet.setColumnWidths(7, 1, 110);
  if (priority_item_keys.length > 0){
    diff_sheet.getRange(2, 5, priority_item_keys.length, priority_item_keys[0].length).setValues(priority_item_keys);
    diff_sheet.getRange(2, 6, priority_item_keys.length, priority_item_keys[0].length).setValues(old_priorities);
    diff_sheet.getRange(2, 7, priority_item_keys.length, priority_item_keys[0].length).setValues(new_priorities);
  }
  diff_sheet.setRowHeight(1, 40);
  diff_sheet.getRange(1,1,1,7).setFontSize(14);
  

  return [[note_item_keys, new_notes], [priority_item_keys, new_priorities]];
}


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Push changes to a Zotero item
function update_zot_item(item_key, content){
  var response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + item_key +'?key=' + api_key);
  options = {
  'method' : 'patch',
  'payload' : JSON.stringify({"note" : content}),
  'headers':{"If-Unmodified-Since-Version": response.getHeaders()["last-modified-version"]}} ;
  response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + item_key +'?key=' + api_key, options);
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Setup the dashboard
function dashboard_setup(){
  SpreadsheetApp.getActiveSpreadsheet().insertSheet('Dashboard');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');

  var headers = [["Total Papers in Collection",	"Total Unread Papers",	'Total High" (<= 5) Priority Unread Papers',
  	'Total "Low/for fun" (> 5) Priority Unread Papers', 'Hidden']];
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  sheet.setRowHeight(1, 50);
  sheet.setColumnWidths(1, 2, 150);
  sheet.setColumnWidths(3, 3, 250);
  sheet.getRange("A:Z").setFontFamily('Cambria');
  sheet.getRange("1:2").setBackground('#434343');
  r = sheet.getRange("1:1");
  r.setFontSize(14);
  r.setFontWeight("bold");
  r.setWrap(true);
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.setRowHeight(2, 35);
  sheet.getRange("A2").setFormula("=COUNTA('Paper Notes'!A2:A)");
  sheet.getRange("B2").setFormula("=COUNTIF('Paper Notes'!K2:K, \"=0\")");
  sheet.getRange("C2").setFormula("=countifs('Paper Notes'!K2:K, \"=0\", 'Paper Notes'!H2:H, \"<=5\")");
  sheet.getRange("D2").setFormula("=countifs('Paper Notes'!K2:K, \"=0\", 'Paper Notes'!H2:H, \">5\")");
  r = sheet.getRange("2:2");
  r.setFontSize(14);
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.getRange("3:3").setBackground("#783f04");
  sheet.setRowHeight(3, 5);

  sheet.getRange("A4:D4").mergeAcross();
  r = sheet.getRange("A4");
  r.setBackground('#434343');
  r.setValue("Next Highest Priority Unread Paper:");
  r.setFontSize(14);
  r.setFontWeight("bold");
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.getRange("E4").setFormula("=MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0)");

  sheet.getRange("A5:D5").mergeAcross();
  r = sheet.getRange("A5");
  r.setBackground('#434343');
  r.setFormula("=HYPERLINK(INDEX('Paper Notes'!B:L, MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0), 8), INDEX('Paper Notes'!B:L, MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0), 1))");
  r.setFontFamily("Inconsolata");
  r.setFontSize(14);
  r.setWrap(true);
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.getRange("A6:D6").mergeAcross();
  r = sheet.getRange("A6");
  r.setBackground('#434343');
  r.setFormula("=INDEX('Paper Notes'!B:L, MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0), 3)");
  r.setFontFamily("Inconsolata");
  r.setFontSize(12);
  r.setWrap(true);
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.getRange("A7:D7").mergeAcross();
  r = sheet.getRange("A7");
  r.setBackground('#434343');
  r.setFormula("=INDEX('Paper Notes'!B:L, MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0), 2)");
  r.setFontFamily("Inconsolata");
  r.setFontSize(12);
  r.setWrap(true);
  r.setFontColor("white");
  r.setHorizontalAlignment("center");
  r.setVerticalAlignment("middle");

  sheet.getRange("A8:B9").mergeAcross();
  sheet.getRange("A8").setValue("Abstract:");
  sheet.getRange("C8:D9").mergeAcross();
  sheet.getRange("C8").setValue("Notes:");
  r = sheet.getRange("8:8");
  r.setFontSize(12);
  r.setFontWeight("bold");
  sheet.getRange("A8:D9").setBackground("#cccccc");

  r = sheet.getRange("A9");
  r.setFormula("=INDEX('Paper Notes'!B:L, MATCH(MINIFS('Paper Notes'!H:H, 'Paper Notes'!K:K, \"<1\"), 'Paper Notes'!H:H, 0), 9)");
  r.setFontSize(11);
  r.setFontFamily("Inconsolata");
  r.setWrap(true);

  r = sheet.getRange("C9");
  r.setFontSize(11);
  r.setWrap(true);
  r.setVerticalAlignment("top");

  sheet.setFrozenRows(8);
  sheet.hideColumns(5);
  var maxRows = sheet.getMaxRows(); 
  var lastRow = sheet.getLastRow();
  sheet.deleteRows(lastRow+1, maxRows-lastRow);
  var maxCols = sheet.getMaxColumns();
  var lastCol = sheet.getLastColumn();
  sheet.deleteColumns(lastCol+1, maxCols-lastCol);

}
