/**
 * GmailService.gs
 * Reads Gmail thread content and extracts structured data for Claude.
 */

var MAX_MESSAGE_CHARS = 3000;

/**
 * Returns a readable transcript of the thread for Claude.
 * @param {string} messageId - The active message ID from the add-on event.
 * @returns {{ transcript: string, participants: Array<{name:string,email:string}>, hasIcsAttachment: boolean }}
 */
function getThreadContent(messageId) {
  var message = GmailApp.getMessageById(messageId);
  var thread = message.getThread();
  var messages = thread.getMessages();

  var transcriptParts = [];
  var hasIcsAttachment = false;

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];

    // Check for .ics attachments
    var attachments = msg.getAttachments();
    for (var j = 0; j < attachments.length; j++) {
      var att = attachments[j];
      if (att.getName().toLowerCase().endsWith('.ics') ||
          att.getContentType().indexOf('calendar') !== -1) {
        hasIcsAttachment = true;
      }
    }

    var body = msg.getPlainBody() || stripHtml(msg.getBody());
    if (body.length > MAX_MESSAGE_CHARS) {
      body = body.substring(0, MAX_MESSAGE_CHARS) + '\n[... truncated ...]';
    }

    transcriptParts.push(
      'From: ' + msg.getFrom() + '\n' +
      'To: ' + msg.getTo() + '\n' +
      'Date: ' + msg.getDate().toISOString() + '\n' +
      'Subject: ' + msg.getSubject() + '\n\n' +
      body
    );
  }

  var participants = extractParticipants(messages);
  var transcript = transcriptParts.join('\n\n---\n\n');

  return {
    transcript: transcript,
    participants: participants,
    hasIcsAttachment: hasIcsAttachment,
    subject: messages.length > 0 ? messages[0].getSubject() : ''
  };
}

/**
 * Replies to a specific Gmail message, threading it into the existing conversation.
 * @param {string} messageId
 * @param {string} body - Plain-text reply body.
 * @returns {{ success: boolean }}
 */
function sendReply(messageId, body) {
  var message = GmailApp.getMessageById(messageId);
  message.reply(body);
  return { success: true };
}

// ---------------------------------------------------------------------------
// Label management
// ---------------------------------------------------------------------------

var LABEL_PREFIX = 'Meeting Scheduler';
var LABEL_AWAITING  = LABEL_PREFIX + '/Awaiting Response';
var LABEL_FOLLOWUP  = LABEL_PREFIX + '/Needs Follow-up';
var LABEL_SCHEDULED = LABEL_PREFIX + '/Scheduled';

/**
 * Applies a Meeting Scheduler tracking label to the thread of a given message.
 * Creates the label (and parent) if they don't exist.
 * @param {string} messageId
 * @param {string} labelName - Full label path, e.g. LABEL_AWAITING.
 */
function applyMeetingLabel(messageId, labelName) {
  try {
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      // Ensure parent label exists first
      var parent = GmailApp.getUserLabelByName(LABEL_PREFIX);
      if (!parent) GmailApp.createLabel(LABEL_PREFIX);
      label = GmailApp.createLabel(labelName);
    }
    GmailApp.getMessageById(messageId).getThread().addLabel(label);
  } catch (e) {
    Logger.log('applyMeetingLabel error (' + labelName + '): ' + e.message);
  }
}

/**
 * Removes all Meeting Scheduler tracking labels from the thread of a given message.
 * @param {string} messageId
 */
function clearMeetingLabels(messageId) {
  try {
    var thread = GmailApp.getMessageById(messageId).getThread();
    [LABEL_AWAITING, LABEL_FOLLOWUP, LABEL_SCHEDULED].forEach(function(name) {
      var label = GmailApp.getUserLabelByName(name);
      if (label) thread.removeLabel(label);
    });
  } catch (e) {
    Logger.log('clearMeetingLabels error: ' + e.message);
  }
}

/**
 * Returns the date of the user's most recent sent message in the thread,
 * and how many days have passed since then.
 * @param {string} messageId
 * @returns {{ date: Date|null, daysSince: number }}
 */
function getLastSentInfo(messageId) {
  try {
    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var messages = GmailApp.getMessageById(messageId).getThread().getMessages();
    for (var i = messages.length - 1; i >= 0; i--) {
      if (messages[i].getFrom().toLowerCase().indexOf(userEmail) !== -1) {
        var date = messages[i].getDate();
        var daysSince = Math.floor((new Date() - date) / (1000 * 60 * 60 * 24));
        return { date: date, daysSince: daysSince };
      }
    }
  } catch (e) {
    Logger.log('getLastSentInfo error: ' + e.message);
  }
  return { date: null, daysSince: 0 };
}

/**
 * Returns summary info for threads with a given label.
 * @param {string} labelName
 * @param {number} max
 * @returns {Array<{ subject:string, daysSince:number, threadUrl:string }>}
 */
function getThreadsForLabel(labelName, max) {
  var results = [];
  try {
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) return results;
    var threads = label.getThreads(0, max);
    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var now = new Date();

    for (var i = 0; i < threads.length; i++) {
      var thread = threads[i];
      var messages = thread.getMessages();
      var lastUserDate = null;
      for (var m = messages.length - 1; m >= 0; m--) {
        if (messages[m].getFrom().toLowerCase().indexOf(userEmail) !== -1) {
          lastUserDate = messages[m].getDate();
          break;
        }
      }
      var daysSince = lastUserDate ? Math.floor((now - lastUserDate) / (1000 * 60 * 60 * 24)) : 0;
      results.push({
        subject: thread.getFirstMessageSubject() || '(no subject)',
        daysSince: daysSince,
        threadUrl: 'https://mail.google.com/mail/u/0/#all/' + thread.getId()
      });
    }
  } catch (e) {
    Logger.log('getThreadsForLabel error (' + labelName + '): ' + e.message);
  }
  return results;
}

/**
 * Deduped list of participants from all From/To/CC headers.
 * @param {GmailMessage[]} messages
 * @returns {Array<{name:string, email:string}>}
 */
function extractParticipants(messages) {
  var seen = {};
  var participants = [];

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var headers = [msg.getFrom(), msg.getTo(), msg.getCc()];

    for (var h = 0; h < headers.length; h++) {
      if (!headers[h]) continue;
      var entries = headers[h].split(',');
      for (var e = 0; e < entries.length; e++) {
        var parsed = parseEmailAddress(entries[e].trim());
        if (parsed && !seen[parsed.email.toLowerCase()]) {
          seen[parsed.email.toLowerCase()] = true;
          participants.push(parsed);
        }
      }
    }
  }

  return participants;
}

/**
 * Parses "Name <email>" or bare "email" into {name, email}.
 * @param {string} raw
 * @returns {{name:string, email:string}|null}
 */
function parseEmailAddress(raw) {
  if (!raw) return null;
  var match = raw.match(/^(.*?)\s*<([^>]+)>\s*$/);
  if (match) {
    return { name: match[1].replace(/^"|"$/g, '').trim(), email: match[2].trim() };
  }
  var emailMatch = raw.match(/[\w.+-]+@[\w.-]+\.\w+/);
  if (emailMatch) {
    return { name: emailMatch[0], email: emailMatch[0] };
  }
  return null;
}

/**
 * Strips HTML tags, returns plain text suitable for Claude.
 * @param {string} html
 * @returns {string}
 */
function stripHtml(html) {
  if (!html) return '';
  // Replace block-level tags with newlines
  var text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '');
  // Decode common HTML entities
  text = text
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
  // Collapse excess whitespace
  text = text.replace(/\n{3,}/g, '\n\n').trim();
  return text;
}
