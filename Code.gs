// ════════════════════════════════════════════════════════════════
//  GLEN DON LODGE — Google Apps Script
//  Paste this into Extensions → Apps Script in your Google Sheet.
//
//  Deploy settings:
//    Execute as:  Me
//    Who can access: Anyone
//
//  After any change: Deploy → Manage deployments → edit → New version
// ════════════════════════════════════════════════════════════════

const BOOKINGS_TAB    = 'Bookings';
const SUBSCRIBERS_TAB = 'Subscribers';
const CABIN_NAME      = 'Glen Don Lodge';

function doGet(e) {
  const action = e.parameter.action || 'getAll';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  try {
    let result;
    switch (action) {
      case 'getAll':
        result = { bookings: getRows(ss, BOOKINGS_TAB), subscribers: getRows(ss, SUBSCRIBERS_TAB) };
        break;
      case 'addBooking':
        const booking = JSON.parse(e.parameter.data);
        addRow(ss, BOOKINGS_TAB, booking);
        sendBookingEmail('confirm', booking);
        result = { success: true };
        break;
      case 'updateBooking':
        updateRow(ss, BOOKINGS_TAB, e.parameter.id, 'status', e.parameter.status);
        if (e.parameter.emailData) {
          const type = e.parameter.status === 'approved' ? 'approved' : 'declined';
          sendBookingEmail(type, JSON.parse(e.parameter.emailData));
        }
        result = { success: true };
        break;
      case 'deleteBooking':
        deleteRow(ss, BOOKINGS_TAB, e.parameter.id);
        result = { success: true };
        break;
      case 'addSubscriber':
        addRow(ss, SUBSCRIBERS_TAB, JSON.parse(e.parameter.data));
        result = { success: true };
        break;
      case 'deleteSubscriber':
        deleteRow(ss, SUBSCRIBERS_TAB, e.parameter.id);
        result = { success: true };
        break;
      case 'broadcast':
        const recipients = JSON.parse(e.parameter.recipients);
        sendBroadcast(e.parameter.subject, e.parameter.message, recipients);
        result = { success: true, sent: recipients.length };
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ── Sheet helpers ─────────────────────────────────────────────

function getRows(ss, tabName) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  const tz      = Session.getScriptTimeZone();
  return values.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var val = row[i];
      if (val instanceof Date) {
        obj[h] = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      } else {
        obj[h] = (val !== undefined && val !== null) ? String(val).trim() : '';
      }
    });
    return obj;
  });
}

function addRow(ss, tabName, data) {
  const sheet   = ss.getSheetByName(tabName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; }));
}

function updateRow(ss, tabName, id, field, value) {
  const sheet   = ss.getSheetByName(tabName);
  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol   = headers.indexOf('id');
  const fldCol  = headers.indexOf(field);
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][idCol]).trim() === String(id).trim()) {
      sheet.getRange(i + 1, fldCol + 1).setValue(value);
      break;
    }
  }
}

function deleteRow(ss, tabName, id) {
  const sheet   = ss.getSheetByName(tabName);
  const values  = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol   = headers.indexOf('id');
  for (var i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idCol]).trim() === String(id).trim()) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

// ── Email helpers ─────────────────────────────────────────────

function sendBookingEmail(type, data) {
  if (!data.email) return;
  const code = data.code || '—';

  const templates = {
    confirm: {
      subject: CABIN_NAME + ' — booking request received!',
      body:
        'Hi ' + data.name + ',\n\n' +
        'Thanks for your booking request at ' + CABIN_NAME + '! 🏡\n\n' +
        '📅 Check-in:  ' + data.checkIn  + '\n' +
        '📅 Check-out: ' + data.checkOut + '\n' +
        '🌙 Nights:    ' + data.nights   + '\n' +
        '👥 Guests:    ' + data.guests   + '\n\n' +
        '─────────────────────────────────\n' +
        '🔑 Your booking code: ' + code + '\n' +
        '─────────────────────────────────\n\n' +
        'Keep this code — together with your email address you can\n' +
        'check the status or cancel your booking by visiting the site\n' +
        'and clicking "My Booking".\n\n' +
        'The cabin admin will review your request and be in touch soon.\n\n' +
        '— ' + CABIN_NAME
    },
    approved: {
      subject: '✅ ' + CABIN_NAME + ' — your stay is confirmed!',
      body:
        'Hi ' + data.name + ',\n\n' +
        'Great news — your stay at ' + CABIN_NAME + ' is CONFIRMED! 🎉\n\n' +
        '📅 Check-in:  ' + data.checkIn  + '\n' +
        '📅 Check-out: ' + data.checkOut + '\n' +
        '🌙 Nights:    ' + data.nights   + '\n' +
        '👥 Guests:    ' + data.guests   + '\n\n' +
        'See you there! 🏡\n\n— ' + CABIN_NAME
    },
    declined: {
      subject: CABIN_NAME + ' — about your booking request',
      body:
        'Hi ' + data.name + ',\n\n' +
        'Unfortunately your requested dates (' + data.checkIn + ' → ' + data.checkOut + ') ' +
        'are not available right now.\n\n' +
        'Please check the calendar for open dates and feel free to try again!\n\n' +
        '— ' + CABIN_NAME + ' 🏡'
    }
  };

  const tpl = templates[type];
  if (!tpl) return;
  MailApp.sendEmail({ to: data.email, subject: tpl.subject, body: tpl.body, name: CABIN_NAME });
}

function sendBroadcast(subject, message, recipients) {
  recipients.forEach(function(r) {
    if (!r.email) return;
    MailApp.sendEmail({
      to: r.email,
      subject: subject,
      body: 'Hi ' + r.name + ',\n\n' + message + '\n\n— The ' + CABIN_NAME + ' admin 🏡\n\n' +
            '─────────────────────────────\nReply to this email to be removed from the list.',
      name: CABIN_NAME
    });
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
