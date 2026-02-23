/**
 * AutoProcessor.gs
 * Background auto-processing for "Auto" mode.
 *
 * Install: run setupAutoTrigger() once from the Apps Script editor, or it will
 * be called automatically when the user enables Auto mode in Settings.
 *
 * What it does every 10 minutes:
 *  1. Scans recent unread inbox emails → if an inbound scheduling request is
 *     found and not yet processed, auto-replies with availability and labels
 *     the thread "Awaiting Response".
 *  2. Checks "Awaiting Response" threads → if the other party has replied and
 *     agreement is reached, auto-creates a Google Meet invite.
 */

var AUTO_TRIGGER_HANDLER = 'runAutoProcessor';
var PROCESSED_IDS_KEY    = 'auto_processed_thread_ids';
var MAX_PROCESSED_IDS    = 150;   // fits well within 9 KB PropertiesService limit

// ---------------------------------------------------------------------------
// Trigger management
// ---------------------------------------------------------------------------

/**
 * Installs a time-based trigger that runs runAutoProcessor every 10 minutes.
 * Safe to call repeatedly — removes any existing trigger first.
 */
function setupAutoTrigger() {
  removeAutoTrigger();
  ScriptApp.newTrigger(AUTO_TRIGGER_HANDLER)
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('Auto processor trigger installed (every 10 min).');
}

/**
 * Removes all auto-processor triggers.
 */
function removeAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === AUTO_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ---------------------------------------------------------------------------
// Main runner
// ---------------------------------------------------------------------------

/**
 * Called by the time-based trigger. Uses a lock to prevent overlapping runs.
 */
function runAutoProcessor() {
  var lock = LockService.getUserLock();
  if (!lock.tryLock(5000)) {
    Logger.log('runAutoProcessor: could not acquire lock, skipping.');
    return;
  }
  try {
    var settings = getSettings();
    if (!settings.autoMode) {
      Logger.log('runAutoProcessor: auto mode disabled, skipping.');
      return;
    }
    processNewInboundEmails(settings);
    processAwaitingReplies(settings);
  } catch (e) {
    Logger.log('runAutoProcessor error: ' + e.message + '\n' + e.stack);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Step 1: Process new inbound scheduling emails
// ---------------------------------------------------------------------------

/**
 * Searches for recent unread inbox emails, analyzes each for scheduling intent,
 * and auto-replies with availability for any new inbound request.
 * @param {{ userName: string }} settings
 */
function processNewInboundEmails(settings) {
  var processedIds = getProcessedThreadIds();
  var processed = 0;

  // Search recent unread inbox emails — broad but capped to keep API usage low
  var threads = GmailApp.search('in:inbox is:unread newer_than:2d', 0, 20);

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var tid = thread.getId();
    if (processedIds.indexOf(tid) !== -1) continue;  // already handled

    try {
      var messages = thread.getMessages();
      if (!messages || messages.length === 0) continue;

      var lastMsg = messages[messages.length - 1];
      var messageId = lastMsg.getId();

      var threadContent = getThreadContent(messageId);
      var analysis = analyzeMeetingThread(threadContent);

      // Mark processed regardless of outcome to avoid re-analyzing
      markThreadProcessed(tid);

      if (!analysis.isSchedulingThread) continue;
      if (analysis.meetingStatus !== 'inbound_request') continue;

      // Auto-reply with calendar availability
      var duration = analysis.durationMinutes || 30;
      var freeSlots = getFreeSlots(duration, 3);
      if (freeSlots.length === 0) {
        Logger.log('Auto: no free slots for thread ' + tid + ', skipping.');
        continue;
      }

      var draftBody = draftReplyEmail(
        threadContent, freeSlots, analysis,
        /*manualAvailability=*/ '',
        duration,
        settings.userName || ''
      );

      sendReply(messageId, draftBody);
      applyMeetingLabel(messageId, LABEL_AWAITING);
      processed++;

      Logger.log('Auto: replied to "' + thread.getFirstMessageSubject() + '"');

    } catch (threadErr) {
      Logger.log('processNewInboundEmails: error on thread ' + tid + ' — ' + threadErr.message);
    }
  }

  Logger.log('processNewInboundEmails: auto-replied to ' + processed + ' threads.');
}

// ---------------------------------------------------------------------------
// Step 2: Check awaiting threads for replies and auto-invite
// ---------------------------------------------------------------------------

/**
 * Checks "Awaiting Response" threads. If the other party has replied and
 * agreement is reached, creates a Google Meet invite automatically.
 * @param {{ userName: string }} settings
 */
function processAwaitingReplies(settings) {
  var label = GmailApp.getUserLabelByName(LABEL_AWAITING);
  if (!label) return;

  var threads = label.getThreads(0, 30);
  var userEmail = Session.getActiveUser().getEmail().toLowerCase();
  var resolved = 0;

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    try {
      var messages = thread.getMessages();
      if (!messages || messages.length === 0) continue;

      // Find user's last message index
      var lastUserIndex = -1;
      for (var m = messages.length - 1; m >= 0; m--) {
        if (messages[m].getFrom().toLowerCase().indexOf(userEmail) !== -1) {
          lastUserIndex = m;
          break;
        }
      }

      var lastMsg = messages[messages.length - 1];
      var otherPartyReplied = lastMsg.getFrom().toLowerCase().indexOf(userEmail) === -1
                              && lastUserIndex >= 0
                              && lastUserIndex < messages.length - 1;

      if (!otherPartyReplied) continue;

      // Re-analyze now that there's a new reply
      var messageId = lastMsg.getId();
      var threadContent = getThreadContent(messageId);
      var analysis = analyzeMeetingThread(threadContent);

      // Remove awaiting label regardless
      thread.removeLabel(label);

      if (analysis.meetingStatus === 'agreement_reached'
          && analysis.agreedTime && analysis.agreedTime.startIso) {
        try {
          createMeetEvent(analysis);
          clearMeetingLabels(messageId);
          applyMeetingLabel(messageId, LABEL_SCHEDULED);
          resolved++;
          Logger.log('Auto: invite sent for "' + thread.getFirstMessageSubject() + '"');
        } catch (inviteErr) {
          Logger.log('Auto: invite error for thread ' + thread.getId() + ' — ' + inviteErr.message);
        }
      }
      // If not agreement_reached, the awaiting label is removed and the
      // contextual card will re-evaluate on next open.

    } catch (threadErr) {
      Logger.log('processAwaitingReplies: error on thread ' + thread.getId() + ' — ' + threadErr.message);
    }
  }

  Logger.log('processAwaitingReplies: auto-invited ' + resolved + ' threads.');
}

// ---------------------------------------------------------------------------
// Processed thread ID tracking
// ---------------------------------------------------------------------------

/**
 * Returns the list of already-processed thread IDs.
 * @returns {string[]}
 */
function getProcessedThreadIds() {
  var stored = PropertiesService.getUserProperties().getProperty(PROCESSED_IDS_KEY);
  if (!stored) return [];
  try { return JSON.parse(stored); } catch (e) { return []; }
}

/**
 * Marks a thread ID as processed so it is not re-analyzed on the next run.
 * Keeps only the most recent MAX_PROCESSED_IDS entries.
 * @param {string} threadId
 */
function markThreadProcessed(threadId) {
  var ids = getProcessedThreadIds();
  if (ids.indexOf(threadId) !== -1) return; // already recorded

  ids.unshift(threadId); // most recent first
  if (ids.length > MAX_PROCESSED_IDS) {
    ids = ids.slice(0, MAX_PROCESSED_IDS);
  }
  PropertiesService.getUserProperties().setProperty(PROCESSED_IDS_KEY, JSON.stringify(ids));
}
