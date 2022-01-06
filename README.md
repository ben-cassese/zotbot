# ZotBot
### A Google Sheets/Zotero project for taking and organizing notes on academic papers

My current workflow: Add papers to Zotero. Let the accumulate unread.

My hopefully more common workflow: Actually read the papers and take notes using this slightly prettier interface.

This is a Google Apps Script which uses the Zotero API to pull data from a personal library into a Google Sheets document. It can then sync edits back to that library.  Basic usage looks like:
- Add papers to Zotero
- Open up this sheet, pull in Zotero data
- Assign priorities to the new papers on Paper Notes
- Open the paper from the Dashboard, read the paper while taking notes on the Dashboard
- Sync my notes and priorities with my Zotero library.

This is my first project that uses any kind of spreadsheet scripting! Had fun but was definitely in a little over my head.


## Some Screenshots
The Dashboard, where I'll take notes on whatever is the highest priority unread paper

<img src="https://user-images.githubusercontent.com/10929214/148453615-b8176d08-b71c-44f3-8505-5917e8768f58.png" width="600">

The Paper Notes tab, where you can see/edit all the notes and priorities for all the papers in the collection. This is all imported automatically from Zotero.

<img src="https://user-images.githubusercontent.com/10929214/148458022-093e7303-0f16-4290-a31d-5ba3e2227afe.png" width="900">

The custom menu to pull/push data to Zotero.

<img src="https://user-images.githubusercontent.com/10929214/148458704-9ca1ec61-f67e-4953-a455-b054549a6d23.png" width="900">


### Setup
---
Open a new Google Sheet. Go to Extensions -> Apps Script. Copy in the ZotBot code, replacing the "redacted" fields at the top with the values appropriate to your Zotero library/collection. Save the code, then select Run. A dialogue will appear to review permissions, and it'll be upset that Google hasn't verified the project, but just continue to the project and select Allow. That should be everything you need to do with the Apps Script editor.

Back in the spreadsheet, there should be a new menu called ZotBot. Under that select Setup Everything. It will take a minute or two to setup with time depending on the size of the collection you're importing.

### Usage
---
Although this pulls in a bunch of information about each paper (like title, abstract, etc), you can only change/sync two fields with your Zotero account. The first is a paper's "priority"- how important it is to read it, small numbers being more important. When you add a paper to Zotero, it will assign a priority of 999. You can edit a paper's priority on the Paper Notes tab.

The second is notes about the paper. You can either edit those on the Paper Notes tab, or in the Dashboard. The Dashboard displays summary data of your whole collection as well as info about the highest priority unread paper. It will display title (hyperlinked to whichever link is saved in Zotero for that paper), authors, date, and the abstract. There's also a cell for notes input, and whenever you modify this cell it'll overwrite the appropriate cell in Paper Notes with that text (and insert "##ZotBot##). The string "##ZotBot##' must be present on the Paper Notes tab in order for the notes to successfully be retrieved from Zotero on a later sync.

When you want to sync edits to your Zotero account, go to ZotBot -> Push changes to Zotero. This will take a bit. Synced changes will appear on the Differences tab and will be pushed to your Zotero account.

When you want to pull new papers in from Zotero, go to ZotBot -> Pull from Zotero. Note that this will overwrite the Paper Notes tab with new data fetched from your Zotero account, so sync any edits before pulling new papers (although since this is Google Sheets, check the version history if you need to recover anything).
