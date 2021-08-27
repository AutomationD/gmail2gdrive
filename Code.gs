// Gmail2GDrive
// https://github.com/ahochsteger/gmail2gdrive
var DEBUG = false;
/**
 * Main function that processes Gmail attachments and stores them in Google Drive.
 * Use this as trigger function for periodic execution.
 */
function Gmail2GDrive() {
  if (!GmailApp) return; // Skip script execution if GMail is currently not available (yes this happens from time to time and triggers spam emails!)
  var config = getGmail2GDriveConfig();
  var end, start, runTime;
  start = new Date(); // Start timer

  Logger.log("INFO: Starting mail attachment processing.");
  if (config.globalFilter===undefined) {
    config.globalFilter = "has:attachment -in:trash -in:drafts -in:spam";
  }

  // Iterate over all rules:
  for (var ruleIdx=0; ruleIdx<config.rules.length; ruleIdx++) {
    var rule = config.rules[ruleIdx];
    var gSearchExp  = config.globalFilter + " " + rule.filter + " -label:" + config.processedLabel;
    if (config.newerThan != "") {
      gSearchExp += " newer_than:" + config.newerThan;
    }
    var doArchive = rule.archive == true;
    var doPDF = rule.saveThreadPDF == true;

    // Process all threads matching the search expression:
    var threads = GmailApp.search(gSearchExp);
    Logger.log("INFO:   Processing rule: "+gSearchExp);
    for (var threadIdx=0; threadIdx<threads.length; threadIdx++) {
      var thread = threads[threadIdx];
      end = new Date();
      runTime = (end.getTime() - start.getTime())/1000;
      Logger.log("INFO:     Processing thread: "+thread.getFirstMessageSubject() + " (runtime: " + runTime + "s/" + config.maxRuntime + "s)");
      if (runTime >= config.maxRuntime) {
        Logger.log("WARNING: Self terminating script after " + runTime + "s");
        return;
      }

      
      if (rule.labels) { // Just label the message if this rule is configured
        
        for (var i=0; i<rule.labels.length; i++) {
          labelText = rule.labels[i]
            .replace('%s',thread.getFirstMessageSubject());
          var label = getOrCreateLabel(labelText);
          
          Logger.log("INFO:           Setting Label '" + labelText +"'");
          thread.addLabel(label);
        }
        
      } else {
        
        // Process all messages of a thread:
        var labels = thread.getLabels();        

        var messages = thread.getMessages();
        for (var msgIdx=0; msgIdx<messages.length; msgIdx++) {
          var message = messages[msgIdx];
          processMessage(message, rule, config, labels);
        }
        
        if (doPDF) { // Generate a PDF document of a thread:
          processThreadToPdf(thread, rule, config, labels);
        }

        
        // Mark a thread as processed:
        var label = getOrCreateLabel(config.processedLabel);
        
        if (!DEBUG)
          thread.addLabel(label);

        if (doArchive) { // Archive a thread if required
          Logger.log("INFO:     Archiving thread '" + thread.getFirstMessageSubject() + "' ...");
          if (!DEBUG)
            thread.moveToArchive();
        }
      }
      

      
    }
  }
  end = new Date(); // Stop timer
  runTime = (end.getTime() - start.getTime())/1000;
  Logger.log("INFO: Finished mail attachment processing after " + runTime + "s");
}

/**
 * Returns the label with the given name or creates it if not existing.
 */
function getOrCreateLabel(labelName) {
  var labels = labelName.split("/");
  var label, labelStem = "";

  for (var i=0; i<labels.length; i++) {

    if (labels[i] !== "") {
      labelStem = labelStem + ((i===0) ? "" : "/") + labels[i];
      label = GmailApp.getUserLabelByName(labelStem) ?
                  GmailApp.getUserLabelByName(labelStem) : GmailApp.createLabel(labelStem);
    }
  }

  return label;
}

/**
 * Recursive function to create and return a complete folder path.
 */
function getOrCreateSubFolder(baseFolder,folderArray) {
  if (folderArray.length == 0) {
    return baseFolder;
  }
  var nextFolderName = folderArray.shift();
  var nextFolder = null;
  var folders = baseFolder.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == nextFolderName) {
      nextFolder = folder;
      break;
    }
  }
  if (nextFolder == null) {
    // Folder does not exist - create it.
    nextFolder = baseFolder.createFolder(nextFolderName);
  }
  return getOrCreateSubFolder(nextFolder,folderArray);
}

/**
 * Returns the GDrive folder with the given path.
 */
function getFolderByPath(path) {
  var parts = path.split("/");

  if (parts[0] == '') parts.shift(); // Did path start at root, '/'?

  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < parts.length; i++) {
    var result = folder.getFoldersByName(parts[i]);
    if (result.hasNext()) {
      folder = result.next();
    } else {
      throw new Error( "folder not found." );
    }
  }
  return folder;
}

/**
 * Returns the GDrive folder with the given name or creates it if not existing.
 */
function getOrCreateFolder(folderName) {
  var folder;
  try {
    folder = getFolderByPath(folderName);
  } catch(e) {
    var folderArray = folderName.split("/");
    folder = getOrCreateSubFolder(DriveApp.getRootFolder(), folderArray);
  }
  return folder;
}

/**
 * Returns formatted file/direcotry name.
 */

function formatName(name, message, attachment, config, tags, thread) {
// function formatName(name="", message, attachment, config, tags, thread) {
  var date = message === null ? thread.getLastMessageDate() : message.getDate();
  var subject = "";
  var filename = "email.pdf";

  if (message && attachment) {     
    subject = message.getSubject();
    filename = attachment.getName();
  } else {
    if (thread) {
      subject = thread.getFirstMessageSubject();    
    }
  }
  
  name = name.replace('%s',subject)
      .replace('%f', filename
        .split('.')
        .slice(0, -1)
        .join('.'));
        
  if (tags){
    for (var i = 0; i < tags.length; i++) {
      name = name
        .replace('{{' + tags[i].key + '}}', tags[i].value == undefined ? "" : tags[i].value);        
    }
          
  }
  
  name = Utilities.formatDate(date, config.timezone, name)
    .replace(':', '');
  
  return name;
}

function getTags(rule, labels){ 
  // Parse label key/value and use it in the file rename
      tags = [];
      if (rule.labelsRegexp && rule.filenameTo) {
        var match = false;
        for (var i = 0; i < labels.length; i++) {
          
          var re = new RegExp(rule.labelsRegexp);
          match = labels[i].getName().match(re);
          
          if (match) {
            if (match.groups.key && match.groups.value) {
              tags.push({"key": match.groups.key, "value": match.groups.value});
              Logger.log("INFO:           Found "+ match.groups.key + ": " + match.groups.value);              
            }
          } 
        }
        
        // filename = formatName(filename, message, attachment, config, label)
        // Logger.log("INFO:           Updated filename '" + file.getName() + "' -> '" + filename + "'");
        // file.setName(filename);
      }
      return tags;
}

/**
 * Processes a message
 */
function processMessage(message, rule, config, labels) {
  Logger.log("INFO:       Processing message: "+message.getSubject() + " (" + message.getId() + ")");
  var messageDate = message.getDate();
  var attachments = message.getAttachments();
  for (var i = 0; i < labels.length; i++) {
    Logger.log("INFO:           labels:" + labels[i].getName());
  }

  for (var attIdx=0; attIdx<attachments.length; attIdx++) {
    var tags = []
    var attachment = attachments[attIdx];
    Logger.log("INFO:         Processing attachment: "+attachment.getName());
    
    if (rule.filenameFromRegexp) {
      var re = new RegExp(rule.filenameFromRegexp);
      match = (attachment.getName()).match(re);
      
      if (!match) {
        Logger.log("INFO:           Rejecting file '" + attachment.getName() + " not matching" + rule.filenameFromRegexp);
        continue;
      }
    }

    
    try {
      
      tags = getTags(rule, labels);
      
      // If we couldn't find a label with a match
      if (!tags) {
        Logger.log("INFO:           Rejecting label '" + labels[i].getName() + ". Can't find a match for" + rule.labelsRegexp);
        continue;
      }
      var folderName = formatName(rule.folder, message, attachment, config, tags)      
      
      Logger.log("INFO:  Saving to folder " + folderName);
      var folder = getOrCreateFolder(folderName);
      var file = folder.createFile(attachment);
      if (rule.filenameTo) {
        var filename = rule.filenameTo;
      } else {
        var filename = file.getName();
      }
      
      var label = {};

      
      if (
          rule.filenameFrom && 
          rule.filenameTo && 
          rule.filenameFrom == file.getName()
        ) 
      {
        filename = formatName(filename, message, attachment, config, tags)

        Logger.log("INFO:           Updating matched filename '" + file.getName() + "' -> '" + filename + "'");
        // file.setName(filename);
      }
      
      else if (rule.filenameTo) {
        filename = formatName(rule.filenameTo,message, attachment, config, tags);
        
        Logger.log("INFO:           Updating filename '" + file.getName() + "' -> '" + filename + "'");
        // file.setName(filename);
      }

      Logger.log("INFO:           Renaming file '" + file.getName() + "' -> '" + filename + "'");
      file.setName(filename);
      

      file.setDescription("Mail title: " + message.getSubject() + "\nMail date: " + message.getDate() + "\nMail link: https://mail.google.com/mail/u/0/#inbox/" + message.getId());

      
      Utilities.sleep(config.sleepTime);
    } catch (e) {
      Logger.log(e);
    }
  }
}

/**
 * Generate HTML code for one message of a thread.
 */
function processThreadToHtml(thread) {
  Logger.log("INFO:   Generating HTML code of thread '" + thread.getFirstMessageSubject() + "'");
  var messages = thread.getMessages();
  var html = "";
  for (var msgIdx=0; msgIdx<messages.length; msgIdx++) {
    var message = messages[msgIdx];
    html += "From: " + message.getFrom() + "<br />\n";
    html += "To: " + message.getTo() + "<br />\n";
    html += "Date: " + message.getDate() + "<br />\n";
    html += "Subject: " + message.getSubject() + "<br />\n";
    html += "<hr />\n";
    html += message.getBody() + "\n";
    html += "<hr />\n";
  }
  return html;
}

/**
* Generate a PDF document for the whole thread using HTML from .
 */
function processThreadToPdf(thread, rule, config, labels) {  
  var tags = getTags(rule, labels);
  var folderName = formatName(rule.folder, null, null, config, tags, thread);
  var folder = getOrCreateFolder(folderName);
  var html = processThreadToHtml(thread);
  var blob = Utilities.newBlob(html, 'text/html');
  
  var filename = formatName(rule.filenameTo, null, null, config, tags, thread);
  Logger.log("INFO: Saving PDF copy of thread to '"+folderName + " as " + filename + "'");
  var pdf = folder.createFile(blob.getAs('application/pdf')).setName(filename);
  return pdf;
}
