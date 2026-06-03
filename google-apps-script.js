// =====================================================
// Google Apps Script — Advisor Calendar Backend
// =====================================================
// Deploy as Web App: Execute as "Me", Access "Anyone"
//
// Sheet tabs required:
//   1. "Events"   — columns: id, name, date, dateLabel, category, description, location
//   2. "SignUps"  — columns: id, eventId, name, email, company, note, timestamp
//
// Security: All POST requests require a token parameter.
// Set your token below:
const SECRET_TOKEN = 'devrev-advisor-2026';

// =====================================================
// GET handler — read events and sign-ups
// =====================================================
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getEvents') {
    return jsonResponse(getEvents());
  }

  if (action === 'getSignups') {
    return jsonResponse(getSignups());
  }

  return jsonResponse({ error: 'Invalid action' });
}

// =====================================================
// POST handler — write events and sign-ups
// =====================================================
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ error: 'Invalid JSON body' });
  }

  const token = body.token;
  if (token !== SECRET_TOKEN) {
    return jsonResponse({ error: 'Unauthorized' });
  }

  const action = body.action;

  switch (action) {
    case 'addEvent':
      return jsonResponse(addEvent(body.event));
    case 'updateEvent':
      return jsonResponse(updateEvent(body.event));
    case 'deleteEvent':
      return jsonResponse(deleteEvent(body.eventId));
    case 'addSignup':
      return jsonResponse(addSignup(body.signup));
    case 'deleteSignup':
      return jsonResponse(deleteSignup(body.signupId));
    default:
      return jsonResponse({ error: 'Invalid action' });
  }
}

// =====================================================
// EVENTS CRUD
// =====================================================
function getEvents() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers

  const headers = data[0];
  const events = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue; // Skip empty rows
    const event = {};
    headers.forEach((header, idx) => {
      event[header] = row[idx] || '';
    });
    events.push(event);
  }

  return events;
}

function addEvent(event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) return { error: 'Events sheet not found' };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => event[h] || '');

  sheet.appendRow(row);
  return { success: true, event: event };
}

function updateEvent(event) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) return { error: 'Events sheet not found' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === event.id) {
      const row = headers.map(h => event[h] || '');
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return { success: true, event: event };
    }
  }

  return { error: 'Event not found' };
}

function deleteEvent(eventId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  if (!sheet) return { error: 'Events sheet not found' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === eventId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { error: 'Event not found' };
}

// =====================================================
// SIGNUPS CRUD
// =====================================================
function getSignups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SignUps');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const signups = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const signup = {};
    headers.forEach((header, idx) => {
      signup[header] = row[idx] || '';
    });
    signups.push(signup);
  }

  return signups;
}

function addSignup(signup) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SignUps');
  if (!sheet) return { error: 'SignUps sheet not found' };

  // Add timestamp
  signup.timestamp = new Date().toISOString();

  // Generate ID if not provided
  if (!signup.id) {
    signup.id = 'su_' + Date.now();
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => signup[h] || '');

  sheet.appendRow(row);
  return { success: true, signup: signup };
}

function deleteSignup(signupId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SignUps');
  if (!sheet) return { error: 'SignUps sheet not found' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === signupId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { error: 'Signup not found' };
}

// =====================================================
// Utility
// =====================================================
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// SETUP — Run this once to create the sheets with headers
// =====================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Events sheet
  let eventsSheet = ss.getSheetByName('Events');
  if (!eventsSheet) {
    eventsSheet = ss.insertSheet('Events');
  }
  eventsSheet.getRange(1, 1, 1, 7).setValues([['id', 'name', 'date', 'dateLabel', 'category', 'description', 'location']]);

  // SignUps sheet
  let signupsSheet = ss.getSheetByName('SignUps');
  if (!signupsSheet) {
    signupsSheet = ss.insertSheet('SignUps');
  }
  signupsSheet.getRange(1, 1, 1, 7).setValues([['id', 'eventId', 'name', 'email', 'company', 'note', 'timestamp']]);

  Logger.log('Sheets created successfully!');
}

// =====================================================
// SEED — Run this once to populate initial events
// =====================================================
function seedEvents() {
  const events = [
    { id: "evt-4", name: "CAB Track Session - CIO", date: "2026-07-14", dateLabel: "", category: "track-session", description: "60 min virtual. CIO perspectives on enterprise AI adoption, governance, and strategic roadmap priorities.", location: "Virtual" },
    { id: "evt-3", name: "Executive Dinner - Atlanta", date: "2026-07-07", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Atlanta, GA" },
    { id: "evt-1", name: "Executive Dinner - Miami", date: "2026-07-15", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Miami, FL" },
    { id: "evt-2", name: "Executive Dinner - Charlotte", date: "2026-07-23", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Charlotte, NC" },
    { id: "evt-5", name: "Advisor Spotlight - July", date: "2026-07-31", dateLabel: "Monthly", category: "spotlight", description: "DevRev media (podcast, interview, social media) on your specific POV on AI-first enterprise transformation.", location: "Virtual (async)" },
    { id: "evt-8", name: "CAB Track Session - Customer Experience", date: "2026-08-11", dateLabel: "", category: "track-session", description: "60 min virtual. CX transformation through AI - support automation, customer intelligence, and proactive engagement strategies.", location: "Virtual" },
    { id: "evt-6", name: "Executive Dinner - Boston", date: "2026-08-04", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Boston, MA" },
    { id: "evt-7", name: "Executive Dinner - NYC", date: "2026-08-12", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "New York, NY" },
    { id: "evt-10", name: "Executive Dinner - Chicago", date: "2026-08-20", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Chicago, IL" },
    { id: "evt-9", name: "Advisor Spotlight - August", date: "2026-08-31", dateLabel: "Monthly", category: "spotlight", description: "DevRev media (podcast, interview, social media) on your specific POV on how companies are rethinking customer experience.", location: "Virtual (async)" },
    { id: "evt-16", name: "CAB Track Session - Customer Experience", date: "2026-09-15", dateLabel: "", category: "track-session", description: "60 min virtual. CX transformation through AI - support automation, customer intelligence, and next-gen engagement.", location: "Virtual" },
    { id: "evt-11", name: "Executive Dinner - Austin", date: "2026-09-08", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Austin, TX" },
    { id: "evt-13", name: "Executive Dinner - San Francisco", date: "2026-09-16", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "San Francisco, CA" },
    { id: "evt-14", name: "Executive Dinner - Seattle", date: "2026-09-24", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Seattle, WA" },
    { id: "evt-17", name: "Advisor Spotlight - September", date: "2026-09-30", dateLabel: "Monthly", category: "spotlight", description: "DevRev media (podcast, interview, social media) on your specific POV on the future of AI in enterprise.", location: "Virtual (async)" },
    { id: "evt-21", name: "Annual Advisor Day", date: "2026-10-13", dateLabel: "Oct 13-15, 2026", category: "track-session", description: "One and a half day, in-person with DevRev leadership. An exclusive gathering to shape what's next - think strategy, product vision, and real talk with the people building it.", location: "TBD" },
    { id: "evt-27", name: "CAB Track Session - AI Infrastructure", date: "2026-11-10", dateLabel: "", category: "track-session", description: "60 min virtual. AI infrastructure maturity - what's scaling, what's breaking, and how to future-proof.", location: "Virtual" },
    { id: "evt-24", name: "Executive Dinner - SoCal", date: "2026-11-03", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Los Angeles / San Diego, CA" },
    { id: "evt-25", name: "Executive Dinner - Atlanta", date: "2026-11-11", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Atlanta, GA" },
    { id: "evt-26", name: "Executive Dinner - Dallas", date: "2026-11-19", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Dallas, TX" },
    { id: "evt-28", name: "Advisor Spotlight - November", date: "2026-11-30", dateLabel: "Monthly", category: "spotlight", description: "DevRev media (podcast, interview, social media) on your specific POV on AI transformation.", location: "Virtual (async)" },
    { id: "evt-30", name: "CAB Track Session - CIO", date: "2026-12-08", dateLabel: "", category: "track-session", description: "60 min virtual. CIO year-end review - what worked, what's on the 2027 roadmap, and strategic priorities ahead.", location: "Virtual" },
    { id: "evt-29", name: "Executive Dinner - NYC", date: "2026-12-01", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "New York, NY" },
    { id: "evt-12", name: "Executive Dinner - Minneapolis", date: "2026-12-09", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Minneapolis, MN" },
    { id: "evt-15", name: "Executive Dinner - Dallas", date: "2026-12-17", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). Curated for senior enterprise leaders.", location: "Dallas, TX" },
    { id: "evt-31", name: "Advisor Spotlight - December", date: "2026-12-31", dateLabel: "Monthly", category: "spotlight", description: "DevRev media (podcast, interview, social media) on your specific POV - key lessons from 2026 and what's next.", location: "Virtual (async)" },
    { id: "evt-32", name: "Executive Dinner - San Francisco", date: "2027-01-13", dateLabel: "", category: "dinner", description: "Intimate dinner (15-20 attendees). AI strategies and priorities for the new year.", location: "San Francisco, CA" },
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Events');
  const headers = ['id', 'name', 'date', 'dateLabel', 'category', 'description', 'location'];

  events.forEach(event => {
    const row = headers.map(h => event[h] || '');
    sheet.appendRow(row);
  });

  Logger.log(`Seeded ${events.length} events.`);
}
