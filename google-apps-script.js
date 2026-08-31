// ============================================================
// CakeHaus Reviews — Google Apps Script (with moderation)
// ============================================================
// SETUP:
// 1. In your Google Sheet, set Row 1 headers to:
//    Timestamp | Name | Rating | Message | Approved
// 2. Paste this code in Extensions > Apps Script
// 3. Deploy > Manage deployments > Edit > New version > Deploy
//
// HOW TO APPROVE REVIEWS:
// - New reviews appear with "Approved" column empty
// - Type YES in the Approved column (column E) for reviews you want shown
// - Only reviews with "YES" in column E appear on the website
// ============================================================

var SHEET_NAME = 'Sheet1';

// Server-side profanity filter (backup layer)
var BLOCKED_WORDS = [
  'fuck','shit','damn','bitch','ass','dick','cock','pussy','cunt',
  'bastard','whore','slut','piss','crap','bollocks','wanker','twat',
  'nigger','nigga','faggot','retard','kak','poes','naaier','doos',
  'moer','fok','msunery','sleg'
];

function containsProfanity(text) {
  var lower = text.toLowerCase().replace(/[^a-z]/g, ' ');
  for (var i = 0; i < BLOCKED_WORDS.length; i++) {
    if (lower.indexOf(BLOCKED_WORDS[i]) !== -1) return true;
  }
  return false;
}

function doGet(e) {
  var params = e.parameter || {};

  // If name + message params exist, this is a review submission
  if (params.name && params.message) {
    // Server-side profanity check
    if (containsProfanity(params.name) || containsProfanity(params.message)) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, reason: 'blocked' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date().toISOString(),
      params.name,
      params.rating || '5',
      params.message,
      ''  // Approved column — empty until you type YES
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Otherwise return only APPROVED reviews
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var reviews = [];

  for (var i = 1; i < data.length; i++) {
    var approved = (data[i][4] || '').toString().trim().toUpperCase();
    if (data[i][1] && approved === 'YES') {
      reviews.push({
        timestamp: data[i][0],
        name: data[i][1],
        rating: data[i][2],
        message: data[i][3]
      });
    }
  }

  var json = JSON.stringify({ reviews: reviews });

  // Support JSONP callback for CORS-free loading
  if (params.callback) {
    return ContentService
      .createTextOutput(params.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
