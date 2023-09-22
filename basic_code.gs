function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
      .addItem('Import Data', 'importData')
      .addItem('Copy to Public', 'copyToPublic')
      .addToUi();
}

function autoRun()
{
  importData();
  copyToPublic();
}

function importData() {
  
  // get the spreadsheet
  var ss = SpreadsheetApp.openById("ID HERE");
  var sheet = ss.getSheets()[0];
  
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).clearContent();
  
  var startRow = 2;
  var startQuery = 0;
  var dataFoundLastAttempt = 0;
  var perPage = 50;
  
  do {
    dataFoundLastAttempt = batchProcess(sheet, startQuery, perPage, startRow);
    startRow += dataFoundLastAttempt;
    startQuery += perPage;
  }
  while (dataFoundLastAttempt > 0);
  
  sheet.getRange(2,1,sheet.getLastRow(),sheet.getLastColumn()).sort([9,10,11]);  
}

function copyToPublic()
{
  var privateDoc = SpreadsheetApp.openById("ID HERE");
  var privateSheet = privateDoc.getSheets()[0];
  var publicDoc = SpreadsheetApp.openById("ID2 HERE");
  var publicSheet = publicDoc.getSheets()[0];
  privateSheet.copyTo(publicDoc);
  publicDoc.deleteSheet(publicSheet);
}

function batchProcess(sheet, startQuery, queryCount, startRow,)
{
  // get all email threads that match label
  var threads = GmailApp.search ("label:fic-to-read", startQuery, queryCount);
  
  if(threads == null || threads.length == 0)
  {
    return 0;
  }
  
  // get all the messages for the current batch of threads
  var messages = GmailApp.getMessagesForThreads (threads);
  
  if(messages == null || messages.length == 0)
  {
    return 0;
  }
  
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
        updateArray.push(parseBody(subject, message.getPlainBody(), message.getDate()));
      }
    }
  }
   
  sheet.getRange(startRow,1,updateArray.length,updateArray[0].length).setValues(updateArray);
  
  return updateArray.length;
}

function parseBody(subject, text, date){
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
  
  var newChapterData = text.match(/(\S*)(.*) posted a new chapter of (.*) \([\d]* words\)/);
  Logger.log("New Chapter Data: " + newChapterData);
  if(newChapterData != null)
  {
    sortTitle = newChapterData[3];
    sortAuthor = newChapterData[1];
  }
  
  var newWork = text.match(/(\S*)(.*) posted a new work/);
  Logger.log("New Work Data: " + newWork);
  if(newWork != null)
  {
    var workTitle = text.match(/(.*) \([\d]* words\)/);
    sortTitle = workTitle[1];
    sortAuthor = newWork[1];
  }
    
  var linkData = text.match(/(https{0,1}:\/\/archiveofourown\.org\/works\/[\d]+)(\/chapters\/[\d]+)*/);
  if(linkData != null)
  {
    chapterLink = linkData[0];
    ficLink = linkData[1];
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
                   
  
  return [displayAuthor, displayTitle, displayChapter, totalChapterCount, complete, andMore, fandoms, date, sortAuthor, sortTitle, sortChapter];
}
