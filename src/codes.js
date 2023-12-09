class GmailProcessor {
  constructor(searchSubject) {
    this.searchQuery = `in:sent subject:(${searchSubject})`;
  }

  searchEmails() {
    return GmailApp.search(this.searchQuery);
  }

  processAllThreads() {
    const threads = this.searchEmails();
    return threads.map(thread => this.processThread(thread));
  }

  processThread(thread) {
    let messages = thread.getMessages();
    let firstMessage = messages[0];
    let recipients = this.extractEmailAddresses(firstMessage.getTo()).join(', ');
    let subject = firstMessage.getSubject();
    let body = firstMessage.getPlainBody();
    let attachments = firstMessage.getAttachments();
    let attachment = attachments.length > 0 ? attachments[0].getName() : '';
    let facultyId = this.extractFacultyId(attachment);

    return [recipients, subject, body, attachment,facultyId];
  }

  sortRecordsByFacultyType(records, fullFacultyPhrase, adjunctFacultyPhrase) {
    let fullFacultyRecords = [];
    let adjunctFacultyRecords = [];

    for (let record of records) {
      let body = record[2]; // Assuming body is the third element in the record array

      if (body.includes(fullFacultyPhrase)) {
        fullFacultyRecords.push(record);
      } else if (body.includes(adjunctFacultyPhrase)) {
        adjunctFacultyRecords.push(record);
      }
    }

    return {
      fullFacultyRecords,
      adjunctFacultyRecords
    };
  }

  extractEmailAddresses(addressInfo) {
    const emailAddresses = [];
    const parts = addressInfo.split(','); // Split by comma to handle multiple recipients
    for (const part of parts) {
      const trimmedPart = part.trim(); // Trim spaces
      const regex = /<([^>]+)>/; // Match email address within angle brackets
      const match = trimmedPart.match(regex);
      if (match) {
        emailAddresses.push(match[1]); // Extracted email address
      } else {
        // Check if trimmedPart is a valid email before adding
        if (/\S+@\S+\.\S+/.test(trimmedPart)) {
          emailAddresses.push(trimmedPart); // Use the part as is
        }
      }
    }
    return emailAddresses;
  }

  extractFacultyId(filename) {
    const match = filename.match(/(^[A-Za-z]\d+)\.zip$/);
    return match ? match[1] : '';
  }
}

class SheetManager {
  constructor() {
    this.sheetCache = {};
  }

  getSheet(sheetName) {
    if (!this.sheetCache[sheetName]) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        throw new Error(`Sheet not found: ${sheetName}`);
      }
      this.sheetCache[sheetName] = sheet;
    }
    return this.sheetCache[sheetName];
  }

  clearContent(sheetName) {
    const sheet = this.getSheet(sheetName);
    if(sheetName === PASS_SHARE_SHEETNAME){
      sheet.getRange("A2:D").clearContent();
    } else {
      sheet.getRange("A2:E").clearContent();
    }
  }

  writeToSheet(sheetName, data) {
    if (data.length > 0) {
      const sheet = this.getSheet(sheetName);
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
  }

  showCompleteMessage(sheetName) {
    Browser.msgBox(`Information has been successfully displayed in ${sheetName}`);
  }
}

/**
 * This function can be used to manually trigger the authorization flow for the script.
 */
function showAuthorizationDialog() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  Browser.msgBox('Authorization has been granted. You can now use the script functionalities.', Browser.Buttons.OK);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Display URL Share Mail Info', 'displayURLShareMailInfo')
    .addSeparator()
    .addItem('Display Result and Pass Mail Info', 'displayResultAndPassMailInfo')
    .addToUi();
}


function displayURLShareMailInfo() {
  try {
    const gmailProcessor = new GmailProcessor(URL_SHARE_SUBJECT);
    let recordsURLShare = gmailProcessor.processAllThreads();
    const sheetManager = new SheetManager();
    sheetManager.clearContent(URL_SHARE_SHEETNAME); // Pass sheet name
    sheetManager.writeToSheet(URL_SHARE_SHEETNAME, recordsURLShare); // Pass sheet name and data
    sheetManager.showCompleteMessage(URL_SHARE_SHEETNAME);
  } catch(e) {
    console.error(`Error while processing URL_SHARE: ${e.toString()} at ${e.stack}`);
    Browser.msgBox(`Error while processing URL_SHARE. Contact the owner of this script.`);
  }
}

function displayResultAndPassMailInfo() {
  try {
    const resultGmailProcessor = new GmailProcessor(RESULT_SHARE_SUBJECT);
    let recordsResultShare = resultGmailProcessor.processAllThreads();
    let sortedRecordsByFacultyType = resultGmailProcessor.sortRecordsByFacultyType(recordsResultShare, RESULT_SHARE_FULFAC_BODY, RESULT_SHARE_ADJFAC_BODY);

    const sheetManager = new SheetManager();

    // Full Faculty Sheet Operations
    sheetManager.clearContent(RESULT_SHARE_FULFAC_SHEETNAME); // Clear existing content
    sheetManager.writeToSheet(RESULT_SHARE_FULFAC_SHEETNAME, sortedRecordsByFacultyType.fullFacultyRecords);
    sheetManager.showCompleteMessage(RESULT_SHARE_FULFAC_SHEETNAME);

    // Adjunct Faculty Sheet Operations
    sheetManager.clearContent(RESULT_SHARE_ADJFAC_SHEETNAME); // Clear existing content
    sheetManager.writeToSheet(RESULT_SHARE_ADJFAC_SHEETNAME, sortedRecordsByFacultyType.adjunctFacultyRecords);
    sheetManager.showCompleteMessage(RESULT_SHARE_ADJFAC_SHEETNAME);

    // Pass Mail Processor and Sheet Operations
    const passGmailProcessor = new GmailProcessor(PASS_SHARE_SUBJECT);
    let recordsPassShare = passGmailProcessor.processAllThreads();
    sheetManager.clearContent(PASS_SHARE_SHEETNAME); // Clear existing content
    sheetManager.writeToSheet(PASS_SHARE_SHEETNAME, recordsPassShare);
    sheetManager.showCompleteMessage(PASS_SHARE_SHEETNAME);

  } catch(e) {
    console.error(`Error while processing RESULT_SHARE and PASS_SHARE: ${e.toString()} at ${e.stack}`);
    Browser.msgBox(`Error while processing RESULT_SHARE and PASS_SHARE. Contact the owner of this script.`);
  }
}
