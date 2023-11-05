function CategorizeFicEmails() 
{
  var dataFoundLastAttempt = 0;
  var perPage = 50;
  
  do 
  {
    dataFoundLastAttempt = batchProcess(perPage);
  }
  while (dataFoundLastAttempt > 0);
  
}

function batchProcess(queryCount)
{
  Logger.log("Starting search batch");

  // get all email threads that match label
  var threads = GmailApp.search ("label:fic-to-read-fic-to-process ", 0, queryCount);
  
  Logger.log("Search batch returned.");

  if(threads == null || threads.length == 0)
  {
    Logger.log("None found.");
    return [];
  }

  Logger.log("Get labels.");

  var toProcessLabel = GmailApp.getUserLabelByName("fic to read/fic to process");
  var completeLabel = GmailApp.getUserLabelByName("fic to read/complete");
  var incompleteLabel = GmailApp.getUserLabelByName("fic to read/WIP");
  var commentLabel = GmailApp.getUserLabelByName("fic to read/comment");
  var unknownLabel = GmailApp.getUserLabelByName("fic to read/unknown");

  Logger.log("Starting Processing Threads.");

  for(var i = 0; i < threads.length; i++)
  {
    var thread = threads[i];
    var status = threadStatus(thread);
    if(status == "complete")
    {
      thread.addLabel(completeLabel);
    } else if (status == "wip") {
      thread.addLabel(incompleteLabel);
    } else if (status == "comment") {
      thread.addLabel(commentLabel);
    } else {
      thread.addLabel(unknownLabel);
    }

    toProcessLabel.removeFromThread(thread);
  }

  return threads.length;
}

function threadStatus(thread){
  var messages = GmailApp.getMessagesForThread(thread);

  if(messages == null || messages.length == 0)
  {
    Logger.log("None found.");
    return unknown;
  }

  for(var i = 0; i < messages.length; i++)
  {
    var message = messages[i];
    var subject = message.getSubject();

    if (subject.includes("comment") || subject.includes("Comment")){
      return "comment"
    } else if(subject.includes("posted")){
      if(messageIsComplete(message.getPlainBody())){
        return "complete";
      }
    } else {
      return "unknown";
    }
  }

  return "wip";
}

function messageIsComplete(text){
  var sortChapter = "";
  var totalChapterCount = "";
  var complete = false;

  var chapterCountData = text.match(/Chapters: ([\d]+)\/([\d]+|\?)/);
  if(chapterCountData != null)
  {
    sortChapter = chapterCountData[1];
    totalChapterCount = chapterCountData[2];
    complete = (sortChapter == totalChapterCount);
  }
                   
  return complete;
}
