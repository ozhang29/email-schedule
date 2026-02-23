/**
 * CalendarService.gs
 * Checks for existing calendar invites and creates Google Meet events.
 */

var CALENDAR_API_BASE = 'https://www.googleapis.com/calendar/v3/calendars/primary/events';

/**
 * Layered detection for whether a calendar invite already exists.
 * @param {MeetingAnalysis} analysis
 * @returns {{ alreadyScheduled: boolean, eventUrl: string|null, method: string }}
 */
function checkExistingInvite(analysis) {
  // Layer 1: Claude's signal + ICS attachment (cheapest)
  if (analysis.calendarInviteSent || analysis.hasIcsAttachment) {
    return { alreadyScheduled: true, eventUrl: null, method: 'claude_signal' };
  }

  // Layers 2 & 3 only make sense if we have an agreed time
  if (!analysis.agreedTime || !analysis.agreedTime.startIso) {
    return { alreadyScheduled: false, eventUrl: null, method: 'no_agreed_time' };
  }

  var agreedStart = new Date(analysis.agreedTime.startIso);
  var windowMs = 60 * 60 * 1000; // ±1 hour

  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var windowStart = new Date(agreedStart.getTime() - windowMs);
    var windowEnd = new Date(agreedStart.getTime() + windowMs);
    var events = calendar.getEvents(windowStart, windowEnd);

    var participantEmails = (analysis.participantEmails || []).map(function(e) {
      return e.toLowerCase();
    });

    // Layer 2: Match by participant email overlap (~95% reliable)
    for (var i = 0; i < events.length; i++) {
      var evt = events[i];
      var guests = evt.getGuestList();
      for (var g = 0; g < guests.length; g++) {
        var guestEmail = guests[g].getEmail().toLowerCase();
        if (participantEmails.indexOf(guestEmail) !== -1) {
          return {
            alreadyScheduled: true,
            eventUrl: evt.getId() ? buildEventUrl(evt.getId()) : null,
            method: 'calendar_participant_match'
          };
        }
      }
    }

    // Layer 3: Title keyword search on same day (~70% reliable)
    if (analysis.meetingTitle) {
      var titleWords = analysis.meetingTitle.toLowerCase().split(/\s+/).filter(function(w) {
        return w.length > 3;
      });
      for (var i = 0; i < events.length; i++) {
        var evt = events[i];
        var evtTitle = evt.getTitle().toLowerCase();
        var matchCount = 0;
        for (var t = 0; t < titleWords.length; t++) {
          if (evtTitle.indexOf(titleWords[t]) !== -1) matchCount++;
        }
        if (matchCount >= Math.ceil(titleWords.length / 2)) {
          return {
            alreadyScheduled: true,
            eventUrl: evt.getId() ? buildEventUrl(evt.getId()) : null,
            method: 'calendar_title_match'
          };
        }
      }
    }
  } catch (e) {
    Logger.log('CalendarService.checkExistingInvite error: ' + e.message);
    // Non-fatal: fall through and let user decide
  }

  return { alreadyScheduled: false, eventUrl: null, method: 'not_found' };
}

/**
 * Creates a Google Calendar event with a Meet link.
 * @param {MeetingAnalysis} analysis
 * @returns {{ eventId: string, eventUrl: string, meetUrl: string, displayTime: string }}
 */
function createMeetEvent(analysis) {
  if (!analysis.agreedTime || !analysis.agreedTime.startIso) {
    throw new Error('No agreed time available to create event.');
  }

  var start = new Date(analysis.agreedTime.startIso);
  var end = new Date(analysis.agreedTime.endIso || analysis.agreedTime.startIso);

  // If endIso is missing or same as start, compute from durationMinutes
  if (end <= start) {
    var duration = analysis.durationMinutes || 60;
    end = new Date(start.getTime() + duration * 60 * 1000);
  }

  var title = analysis.meetingTitle || 'Meeting';
  var guestEmails = (analysis.participantEmails || []).join(',');
  var options = {
    guests: guestEmails,
    sendInvites: true
  };

  // Step 1: Create event via CalendarApp
  var calEvent = CalendarApp.getDefaultCalendar().createEvent(title, start, end, options);
  var eventId = calEvent.getId();

  // Step 2: PATCH via REST API to add Google Meet conference link
  //  CalendarApp doesn't expose conferenceDataVersion, so we must use REST.
  var meetUrl = addMeetLinkViaRest(eventId, title, start, end);

  return {
    eventId: eventId,
    eventUrl: buildEventUrl(eventId),
    meetUrl: meetUrl,
    displayTime: analysis.agreedTime.displayText || start.toLocaleString()
  };
}

/**
 * PATCHes a Calendar event via REST API to attach a Google Meet conference.
 * @param {string} eventId
 * @param {string} title
 * @param {Date} start
 * @param {Date} end
 * @returns {string} The Meet URL, or empty string if unavailable.
 */
function addMeetLinkViaRest(eventId, title, start, end) {
  // Strip the @google.com suffix CalendarApp adds to event IDs
  var cleanId = eventId.replace(/@.*$/, '');

  var requestId = Utilities.getUuid();
  var patchBody = {
    summary: title,
    conferenceData: {
      createRequest: {
        requestId: requestId,
        conferenceSolutionKey: { type: 'hangoutsMeet' }
      }
    }
  };

  var url = CALENDAR_API_BASE + '/' + encodeURIComponent(cleanId) + '?conferenceDataVersion=1';
  var options = {
    method: 'patch',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(patchBody),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    Logger.log('Calendar REST PATCH failed ' + responseCode + ': ' + response.getContentText().substring(0, 300));
    return '';
  }

  var responseJson = JSON.parse(response.getContentText());
  try {
    var entryPoints = responseJson.conferenceData.entryPoints;
    for (var i = 0; i < entryPoints.length; i++) {
      if (entryPoints[i].entryPointType === 'video') {
        return entryPoints[i].uri;
      }
    }
  } catch (e) {
    Logger.log('Could not extract Meet URL from response: ' + e.message);
  }
  return '';
}

/**
 * Returns up to numSlots free calendar slots over the next 14 weekdays.
 * Slots are within 9am–5pm in the calendar's default timezone.
 * @param {number} durationMinutes
 * @param {number} numSlots
 * @returns {Array<{ startIso: string, endIso: string, displayText: string }>}
 */
function getFreeSlots(durationMinutes, numSlots) {
  var duration = (durationMinutes || 30) * 60 * 1000; // ms — minimum gap required to list a window
  var slots = [];
  var calendar = CalendarApp.getDefaultCalendar();

  var now = new Date();
  // Skip today if current ET time is after 3pm (15:00)
  var etHour = parseInt(Utilities.formatDate(now, 'America/New_York', 'H'), 10);
  var startDayOffset = (etHour >= 15) ? 1 : 0;

  for (var dayOffset = startDayOffset; dayOffset < 21 && slots.length < numSlots; dayOffset++) {
    var day = new Date(now);
    day.setDate(day.getDate() + dayOffset);

    // Skip weekends (0 = Sunday, 6 = Saturday)
    var dow = day.getDay();
    if (dow === 0 || dow === 6) continue;

    // Build 9am and 5pm boundaries in ET for this day
    var dateStr = Utilities.formatDate(day, 'America/New_York', 'yyyy-MM-dd');
    var dayStart = new Date(dateStr + 'T09:00:00-05:00');
    var dayEnd   = new Date(dateStr + 'T17:00:00-05:00');

    // Adjust for DST: use the actual ET offset on that day
    // Recalculate using the formatted offset from the calendar timezone
    var etOffset = Utilities.formatDate(day, 'America/New_York', 'Z'); // e.g. "-0500" or "-0400"
    var sign = etOffset.charAt(0) === '-' ? -1 : 1;
    var offsetHours = parseInt(etOffset.substring(1, 3), 10);
    var offsetMins  = parseInt(etOffset.substring(3, 5), 10);
    var offsetMs = sign * (offsetHours * 60 + offsetMins) * 60 * 1000;
    dayStart = new Date(Date.UTC(
      parseInt(dateStr.substring(0, 4), 10),
      parseInt(dateStr.substring(5, 7), 10) - 1,
      parseInt(dateStr.substring(8, 10), 10),
      9, 0, 0
    ) - offsetMs);
    dayEnd = new Date(Date.UTC(
      parseInt(dateStr.substring(0, 4), 10),
      parseInt(dateStr.substring(5, 7), 10) - 1,
      parseInt(dateStr.substring(8, 10), 10),
      17, 0, 0
    ) - offsetMs);

    // Don't look in the past
    if (dayEnd <= now) continue;
    var searchStart = dayStart < now ? now : dayStart;

    var events = calendar.getEvents(dayStart, dayEnd);
    // Sort events by start time
    events.sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });

    // Find free gaps and report them as availability windows (up to 3 hours each)
    var MAX_BLOCK_MS = 3 * 60 * 60 * 1000; // cap display window at 3 hours
    var cursor = searchStart;
    for (var i = 0; i <= events.length && slots.length < numSlots; i++) {
      var gapEnd = (i < events.length) ? events[i].getStartTime() : dayEnd;
      if (gapEnd - cursor >= duration) {
        var slotStart = new Date(cursor);
        // Show the full free window, capped at 3 hours and the end of the work day
        var slotEnd = new Date(Math.min(cursor.getTime() + MAX_BLOCK_MS, gapEnd.getTime(), dayEnd.getTime()));
        slots.push({
          startIso: slotStart.toISOString(),
          endIso:   slotEnd.toISOString(),
          displayText: Utilities.formatDate(slotStart, 'America/New_York', 'EEE, MMM d') +
                       ', ' +
                       Utilities.formatDate(slotStart, 'America/New_York', 'h:mm a') +
                       ' – ' +
                       Utilities.formatDate(slotEnd, 'America/New_York', 'h:mm a z')
        });
      }
      if (i < events.length) {
        var evtEnd = events[i].getEndTime();
        if (evtEnd > cursor) cursor = evtEnd;
      }
    }
  }

  return slots;
}

/**
 * Checks each proposed time against the user's primary calendar.
 * Only times with a parseable startIso are checked; others get free:null.
 * @param {Array<{ proposedBy:string, displayText:string, startIso:string|null }>} proposedTimes
 * @param {number} durationMinutes
 * @returns {Array<{ proposedBy:string, displayText:string, startIso:string|null,
 *                   endIso:string|null, free:boolean|null, availLabel:string }>}
 */
function checkProposedTimes(proposedTimes, durationMinutes) {
  var duration = (durationMinutes || 30) * 60 * 1000;
  var calendar = CalendarApp.getDefaultCalendar();

  return (proposedTimes || []).map(function(pt) {
    if (!pt.startIso || !pt.startIso.match(/^\d{4}-\d{2}-\d{2}T/)) {
      return { proposedBy: pt.proposedBy || '', displayText: pt.displayText || '', startIso: null, endIso: null, free: null, availLabel: '' };
    }
    try {
      var start = new Date(pt.startIso);
      if (isNaN(start.getTime())) {
        return { proposedBy: pt.proposedBy || '', displayText: pt.displayText || pt.startIso, startIso: pt.startIso, endIso: null, free: null, availLabel: '' };
      }
      var end = new Date(start.getTime() + duration);
      var events = calendar.getEvents(start, end).filter(function(evt) {
        if (evt.isAllDayEvent()) return false;
        var status = evt.getMyStatus();
        return status !== CalendarApp.GuestStatus.NO;
      });
      var free = events.length === 0;
      return {
        proposedBy: pt.proposedBy || '',
        displayText: pt.displayText || pt.startIso,
        startIso: pt.startIso,
        endIso: end.toISOString(),
        free: free,
        availLabel: free ? '✓ Free' : '✗ Busy'
      };
    } catch (e) {
      Logger.log('checkProposedTimes error for ' + pt.startIso + ': ' + e.message);
      return { proposedBy: pt.proposedBy || '', displayText: pt.displayText || pt.startIso, startIso: pt.startIso, endIso: null, free: null, availLabel: '' };
    }
  });
}

/**
 * Builds a Google Calendar event URL from an event ID.
 * @param {string} eventId
 * @returns {string}
 */
function buildEventUrl(eventId) {
  var cleanId = eventId.replace(/@.*$/, '');
  return 'https://calendar.google.com/calendar/event?eid=' + Utilities.base64Encode(cleanId + ' primary');
}
