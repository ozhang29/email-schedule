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
