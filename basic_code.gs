function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
      .addItem('Import Data', 'importData')
      .addItem('Import Lots of Data', 'startImportDataSplitComp')
      .addToUi();
}

function autoRun()
{
  resetProperties();
  importDataSplitComp();
}

//This starts failing reliably when attempting to process over ~1900 emails
function importData()
{
  // get the spreadsheet
  var ss = SpreadsheetApp.openById("ID HERE");
  var sheet = ss.getSheets()[0];
  
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
  
  var allData = [];
  var startRow = 2;
  var startQuery = 0;
  var dataFoundLastAttempt = [];
  var perPage = 200;
  
  do
  {
    dataFoundLastAttempt = batchProcess(startQuery, perPage);
    startRow += dataFoundLastAttempt.length;
    startQuery += perPage;

    allData = allData.concat(dataFoundLastAttempt);
    Logger.log("Found: " + dataFoundLastAttempt.length + " new rows in most recent batch.");
    Logger.log("Found: " + allData.length + " entries so far.");
  }
  while (dataFoundLastAttempt.length > 0);

  sheet.getRange(2,1,allData.length,allData[0].length).setValues(allData);
  
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).sort([9,10,11]);  
}

function resetProperties()
{
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('computationRunning', false);
  userProperties.setProperty('startRow', 2);
  userProperties.setProperty('startQuery', 0);
}

function startImportDataSplitComp()
{
  resetProperties();
  importDataSplitComp();
}

//if the above version frequently times out this version is designed to split computation over multiple executions
function importDataSplitComp(e) 
{
  if(e)
  {
    deleteTriggerById(e.triggerUid);
  }
  var startTime = (new Date()).getTime();

  // get the spreadsheet
  var ss = SpreadsheetApp.openById("ID HERE");
  var sheet = ss.getSheets()[0];

  var startRow = 2;
  var startQuery = 0;
  var dataFoundLastAttempt = [];
  var perPage = 200;

  var userProperties = PropertiesService.getUserProperties();

  if(userProperties.getProperty('computationRunning') == 'true')
  {
    startRow = parseInt(userProperties.getProperty('startRow'));
    startQuery = parseInt(userProperties.getProperty('startQuery'));
    Logger.log("In progress computation found starting from row " + startRow + " query " + startQuery);
  }
  else
  {
    Logger.log("No inprogress computation found clearing sheet");
    sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
    userProperties.setProperty('computationRunning', true);
  }

  do
  {
    dataFoundLastAttempt = batchProcess(startQuery, perPage);

    if(dataFoundLastAttempt.length > 0)
    {
      //Todo: Maybe switch to only writing when approaching timelimit to better utilize time?
      sheet.getRange(startRow,1,dataFoundLastAttempt.length,dataFoundLastAttempt[0].length).setValues(dataFoundLastAttempt);

      startRow += dataFoundLastAttempt.length;
      startQuery += perPage;
      userProperties.setProperty('startRow', startRow);
      userProperties.setProperty('startQuery', startQuery);

      var currentTime = (new Date()).getTime();
      if(currentTime - startTime >= (4 * 60 * 1000)) {
        Logger.log("Taking too long setting trigger to bypass execution limit.");
        ScriptApp.newTrigger('importDataSplitComp').timeBased().after(60 * 1000).create();
        return;
      }
    }
  } while (dataFoundLastAttempt.length > 0);

  Logger.log("Completed checking all messages.");

  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).sort([9,10,11]);

  Logger.log("Sorted Sheet.");

  resetProperties();

  Logger.log("Marked Complete.");

}

function deleteTriggerById(id) {
  let trigger = ScriptApp.getProjectTriggers().find(trigger => trigger.getUniqueId().toString() === id);
    
  if (trigger) {
    Logger.log("Found trigger to delete");
    ScriptApp.deleteTrigger(trigger);
  }
}

function batchProcess(startQuery, queryCount)
{
  Logger.log("Starting search batch");

  // get all email threads that match label
  var threads = GmailApp.search ("label:fic-to-read", startQuery, queryCount);
  
  Logger.log("Search batch returned.");

  if(threads == null || threads.length == 0)
  {
    Logger.log("None found.");
    return [];
  }

  Logger.log("Starting Get Message Batch.");

  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);

  Logger.log("Get Messages Batch competed.");

  if(messages == null || messages.length == 0)
  {
    Logger.log("None found.");
    return [];
  }

  Logger.log("Starting parsing messages.");

  var updateArray = [];
  
  var iMax = Math.min(messages.length, 1000);
  
  for(var i = 0; i < iMax; i++)
  {
    for(var j = 0; j < messages[i].length; j++)
    {
      var message = messages[i][j];
      var subject = message.getSubject();
      if(subject.includes("posted"))
      {
        updateArray.push(parseBody(subject, message.getPlainBody(), message.getDate(), message.getId()));
      }
    }
  }

  Logger.log("Message parsing complete.");

  return updateArray;
}

function parseBody(subject, text, date, messageId){
  var displayTitle = "";
  var sortTitle = ""
  var displayAuthor = "";
  var sortAuthor = "";
  var displayChapter = "";
  var sortChapter = "";
  var chapterLink = "";
  var ficLink = "";
  var authorLink = "";
  var totalChapterCount = "";
  var complete = false;
  var fandoms = "";
  var andMore = false;
  var ficId = "";

  var newChapterData = text.match(/(\S*)(.*) posted a new chapter of (.*) \([\d]* words\)/);
  //Logger.log("New Chapter Data: " + newChapterData);
  if(newChapterData != null)
  {
    sortTitle = newChapterData[3];
    sortAuthor = newChapterData[1];
  }
  
  var newWork = text.match(/(\S*)(.*) posted a new work/);
  //Logger.log("New Work Data: " + newWork);
  if(newWork != null)
  {
    var workTitle = text.match(/(.*) \([\d]* words\)/);
    sortTitle = workTitle[1];
    sortAuthor = newWork[1];
  }
    
  var linkData = text.match(/(https{0,1}:\/\/archiveofourown\.org\/works\/([\d]+))(\/chapters\/[\d]+)*/);
  if(linkData != null)
  {
    chapterLink = linkData[0];
    ficLink = linkData[1];
    ficId = linkData[2];
    displayTitle = '=HYPERLINK("' + ficLink + '","' + sortTitle + '")';
  }
  
  var authorLinkData = text.match(/https{0,1}:\/\/archiveofourown\.org\/users\/.+?\/(pseuds\/[^\)]+)*/);
  if(authorLinkData != null)
  {
    authorLink = authorLinkData[0];
  }
  
  displayAuthor ='=HYPERLINK("' + authorLink + '","' + sortAuthor + '")';

  var chapterCountData = text.match(/Chapters: ([\d]+)\/([\d]+|\?)/);
  if(chapterCountData != null)
  {
    sortChapter = chapterCountData[1];
    displayChapter = '=HYPERLINK("' + chapterLink + '","' + chapterCountData[1] + '")';
    totalChapterCount = chapterCountData[2];
    complete = (sortChapter == totalChapterCount);
  }
  
  var fandomData = text.match(/Fandom: (.*)/);
  if(fandomData != null)
  {
    fandoms = fandomData[1];
    fandoms = fandoms.replace("僕のヒーローアカデミア | Boku no Hero Academia | My Hero Academia", "My Hero Academia");
    fandoms = fandoms.replace("Harry Potter - J. K. Rowling", "Harry Potter");
    fandoms = fandoms.replace("Spider-Man: Into the Spider-Verse (2018)", "Into the Spider-Verse");
  }
  
  var andMoreData = subject.match(/and [\d]+ more/);
  if(andMoreData != null)
  {
    andMore = true;
  }

  return [displayAuthor, displayTitle, displayChapter, totalChapterCount, complete, andMore, fandoms, date, sortAuthor, sortTitle, sortChapter, ficId];
}
