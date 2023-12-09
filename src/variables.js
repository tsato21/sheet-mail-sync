const INITIAL_SETTING_SHEET_NAME = "initial-setting";
const INITIAL_SETTING_SHEET = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(INITIAL_SETTING_SHEET_NAME);

//Declare variables for subject information of target emails
const URL_SHARE_SUBJECT = INITIAL_SETTING_SHEET.getRange("C1").getValue();
const RESULT_SHARE_SUBJECT = INITIAL_SETTING_SHEET.getRange("C2").getValue();
const PASS_SHARE_SUBJECT = INITIAL_SETTING_SHEET.getRange("C3").getValue();
//Declare variables for body information of target emails
const RESULT_SHARE_FULFAC_BODY = INITIAL_SETTING_SHEET.getRange("C4").getValue();
const RESULT_SHARE_ADJFAC_BODY = INITIAL_SETTING_SHEET.getRange("C5").getValue();
//Declare variables for display sheets for target emails
const URL_SHARE_SHEETNAME = INITIAL_SETTING_SHEET.getRange("C6").getValue();
const RESULT_SHARE_FULFAC_SHEETNAME = INITIAL_SETTING_SHEET.getRange("C7").getValue();
const RESULT_SHARE_ADJFAC_SHEETNAME = INITIAL_SETTING_SHEET.getRange("C8").getValue();
const PASS_SHARE_SHEETNAME = INITIAL_SETTING_SHEET.getRange("C9").getValue();