const library_key = redacted;
const collection_key = redacted;
const api_key = redacted;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ZotBot')
      .addItem('Pull data from Zotero', 'fetch_Zot_Data')
      .addToUi();
}

function fetch_Zot_Data() {
  SpreadsheetApp.getActive().toast('Pulling data from Zotero', "ZotBot Test");
  // Set the headers
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GAS Version');
  var headers = [["Zotero Paper Key", "Author",	"Title",
  	"Dates",	"URL", "Abstract",	"Note Zotero Key",	"Notes",	"Priority Zotero Key",	"Priority",	"Read"]]
  sheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);

  // Get the first 100 papers
  var url = 'https://api.zotero.org/users/' + library_key +'/collections/' + collection_key + '/items/top?key=' + api_key + '&limit=100'
  var response = UrlFetchApp.fetch(url);
  var papers = JSON.parse(response.getContentText());

  // Get all of the other papers, if there are more than 100
  var start = 100;
  const num_items = response.getAllHeaders()['total-results'];
  while (num_items > start){
    response = UrlFetchApp.fetch(url + '&start=' + start.toString());
    papers.push(...JSON.parse(response.getContentText()))
    start += 100;
  }

  // Set the columns you can get from the parent item
  var paper_keys = []
  var authors = []
  var titles = []
  var dates = []
  var urls = []
  var abstracts = [];
  for (let i = 0; i < papers.length; i++) {
    paper_keys.push([papers[i]['key']])
    authors.push([papers[i]['meta']['creatorSummary']])
    titles.push([papers[i]['data']['title']])
    dates.push([papers[i]['data']['date']])
    urls.push([papers[i]['data']['url']])
    abstracts.push([papers[i]['data']['abstractNote']])
    }
  sheet.getRange(2, 1, paper_keys.length, paper_keys[0].length).setValues(paper_keys);
  sheet.getRange(2, 2, paper_keys.length, paper_keys[0].length).setValues(authors);
  sheet.getRange(2, 3, paper_keys.length, paper_keys[0].length).setValues(titles);
  sheet.getRange(2, 4, paper_keys.length, paper_keys[0].length).setValues(dates);
  sheet.getRange(2, 5, paper_keys.length, paper_keys[0].length).setValues(urls);
  sheet.getRange(2, 6, paper_keys.length, paper_keys[0].length).setValues(abstracts);

  var note_keys = []
  var notes = []
  var read = []
  var priority_keys = []
  var priorities =[]

  for (let i = 0; i < papers.length; i++) {
    var paper_key = papers[i]['key'];

    var url = 'https://api.zotero.org/users/' + library_key + '/items/' + paper_key + '/children' + '?key=' + api_key
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
          priority_keys.push([priority_key])
          priorities.push([Number(children[i]['data']['note'].split('@@Priority: ')[1].split('@@')[0])])
        }
      }
    }

    // If the file has been added to Zotero since the last time this ran, add the notes
    if (note_key == ''){
      var k = create_ZotBot_notes(paper_key)
      note_key = k[0]
      note_keys.push([note_key])
      notes.push(['##ZotBot##'])
      priority_key = k[1]
      priority_keys.push([priority_key])
      priorities.push([999])
      read.push([0])
    }
  }

  sheet.getRange(2, 7, note_keys.length, note_keys[0].length).setValues(note_keys);
  sheet.getRange(2, 8, note_keys.length, note_keys[0].length).setValues(notes);
  sheet.getRange(2, 9, note_keys.length, note_keys[0].length).setValues(priority_keys);
  sheet.getRange(2, 10, note_keys.length, note_keys[0].length).setValues(priorities);
  sheet.getRange(2, 11, note_keys.length, note_keys[0].length).setValues(read);

  SpreadsheetApp.getActive().toast('Data pull complete', "ZotBot Test");

}



function create_ZotBot_notes(paper_key){
  var response = UrlFetchApp.fetch(url='https://api.zotero.org/items/new?itemType=note')
  const note_template = JSON.parse(response.getContentText());
  var zotbot_note = JSON.parse(JSON.stringify(note_template));
  zotbot_note['note'] = '##ZotBot##';
  var priority_note = JSON.parse(JSON.stringify(note_template));
  priority_note['note'] = '@@Priority: 999@@'


  var headers = {
  "Zotero-Write-Token": Utilities.getUuid().split('-').join(''),
  "Content-Type": "application/json"}

  var options = {
  'method' : 'post',
  'contentType': 'application/json',
  'payload' : JSON.stringify([zotbot_note, priority_note]),
  'headers' : headers}

  var response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items?key=' + api_key, options);
  var written = JSON.parse(response.getContentText())
  var note_key = written["success"][0]
  var priority_key = written["success"][1]

  options = {
  'method' : 'patch',
  'payload' : JSON.stringify({"parentItem" : paper_key}),
  'headers':{"If-Unmodified-Since-Version": response.getHeaders()["last-modified-version"]}
  }
  response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + note_key +'?key=' + api_key, options);
  response = UrlFetchApp.fetch('https://api.zotero.org/users/' + library_key + '/items/' + priority_key +'?key=' + api_key, options);

  return [note_key, priority_key]
}
