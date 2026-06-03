/**
 * ============================================
 * ADVISOR CALENDAR - Google Apps Script Backend
 * ============================================
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire code into Code.gs
 * 3. Create a Google Sheet and note its ID (from the URL)
 * 4. Replace SPREADSHEET_ID below with your sheet's ID
 * 5. Run the setup() function once to create the required sheets
 * 6. Deploy as Web App:
 *    - Click Deploy > New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copy the deployment URL and replace SHEETS_URL in index.html
 *
 * SHEETS CREATED:
 * - "Events" — stores all event data (admin-managed)
 * - "Registrations" — stores all sign-ups (shared across users)
 */

const SPREADSHEET_ID = '1lR4u6FTU9V8bSKHTptUhc4TXSiB083Ka2LpYD3ROkM0'; // <-- Replace this!

// ============================================
// SETUP - Run once to create sheets
// ============================================
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Create Events sheet
  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet('Events');
    eventsSheet.appendRow(['id', 'name', 'date', 'category', 'location', 'description', 'dateLabel']);
  }

  // Create Registrations sheet
  let regsSheet = ss.getSheetByName('Registrations');
  if (!regsSheet) {
    regsSheet = ss.insertSheet('Registrations');
    regsSheet.appendRow(['timestamp', 'eventId', 'eventName', 'eventDate', 'eventMonth', 'name', 'email', 'company', 'note']);
  }

  Logger.log('Setup complete! Sheets created.');
}

// ============================================
// GET handler - returns events or hosts
// ============================================
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getEvents') {
    return getEvents();
  } else if (action === 'getHosts') {
    return getHosts();
  }

  return ContentService
    .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// POST handler - register, add/update/delete events
// ============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    switch (action) {
      case 'register':
        return registerForEvent(data);
      case 'addEvent':
        return addEvent(data);
      case 'updateEvent':
        return updateEvent(data);
      case 'deleteEvent':
        return deleteEvent(data);
      default:
        // Legacy support: if no action, treat as registration
        return registerForEvent(data);
    }
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// GET EVENTS - Returns all events from the Events sheet
// ============================================
function getEvents() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Events');

  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ events: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  const events = data
    .filter(row => row[0]) // Skip empty rows
    .map(row => {
      const event = {
        id: row[0],
        name: row[1],
        date: row[2],
        category: row[3],
        location: row[4],
        description: row[5]
      };
      if (row[6]) event.dateLabel = row[6];
      return event;
    });

  return ContentService
    .createTextOutput(JSON.stringify({ events }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// GET HOSTS - Returns all registrations grouped by eventId
// ============================================
function getHosts() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Registrations');

  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ hosts: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const hosts = {};

  data.forEach(row => {
    if (!row[1]) return; // Skip empty rows
    const eventId = row[1];
    if (!hosts[eventId]) hosts[eventId] = [];
    hosts[eventId].push({
      name: row[5],
      email: row[6],
      company: row[7],
      note: row[8] || ''
    });
  });

  return ContentService
    .createTextOutput(JSON.stringify({ hosts }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// REGISTER - Add a new registration
// ============================================
function registerForEvent(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Registrations');

  sheet.appendRow([
    new Date().toISOString(),
    data.eventId || '',
    data.event || '',
    data.eventDate || '',
    data.eventMonth || '',
    data.name || '',
    data.email || '',
    data.company || '',
    data.note || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// ADD EVENT - Admin adds a new event
// ============================================
function addEvent(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Events');

  sheet.appendRow([
    data.id || 'evt-' + Date.now(),
    data.name || '',
    data.date || '',
    data.category || '',
    data.location || '',
    data.description || '',
    data.dateLabel || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// UPDATE EVENT - Admin edits an existing event
// ============================================
function updateEvent(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Events');
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([[
        data.id,
        data.name || '',
        data.date || '',
        data.category || '',
        data.location || '',
        data.description || '',
        data.dateLabel || ''
      ]]);
      break;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// DELETE EVENT - Admin removes an event
// ============================================
function deleteEvent(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Events');
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
