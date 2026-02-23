/**
 * TriggerService.gs
 * Time-driven background trigger for proactive meeting tracking and follow-up nudges.
 *
 * One-time setup: run setupSchedulerTrigger() from the Apps Script editor to install
 * the daily job. After that, everything runs automatically.
 */

var FOLLOW_UP_AFTER_DAYS = 3;

/**
 * Run this ONCE from the Apps Script editor to install the daily trigger.
 * Go to: Extensions → Apps Script → select this function → Run.
 * Safe to call again — it removes any existing trigger before creating a new one.
 */
function setupSchedulerTrigger() {
  // Remove any existing daily scan triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runDailyMeetingScan') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDailyMeetingScan')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Daily meeting scan trigger installed. Runs every morning at 8am.');
}

/**
 * Daily job: scans threads labeled "Meeting Scheduler/Awaiting Response".
 *
 * For each thread:
 * - If the other party has replied since your last email → remove "Awaiting" label
 *   (the add-on will re-analyze and show the right card when you open it).
 * - If still awaiting and it has been >= FOLLOW_UP_AFTER_DAYS days → add "Needs Follow-up" label.
 *
 * Note: re-analysis uses the AI API. If you want to avoid API calls in the daily job,
 * you can remove the analyzeMeetingThread call and rely on the label removal only.
 */
function runDailyMeetingScan() {
  try {
    var label = GmailApp.getUserLabelByName(LABEL_AWAITING);
    if (!label) {
      Logger.log('No "' + LABEL_AWAITING + '" label found — nothing to scan.');
      return;
    }

    var threads = label.getThreads(0, 30);
    var userEmail = Session.getActiveUser().getEmail().toLowerCase();
    var now = new Date();
    var scanned = 0;
    var nudged = 0;
    var resolved = 0;

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      try {
        scanned++;
        var messages = thread.getMessages();
        if (!messages || messages.length === 0) continue;

        // Find the user's most recent email and check if anyone replied after it
        var lastUserIndex = -1;
        for (var m = messages.length - 1; m >= 0; m--) {
          if (messages[m].getFrom().toLowerCase().indexOf(userEmail) !== -1) {
            lastUserIndex = m;
            break;
          }
        }

        var lastUserDate = lastUserIndex >= 0 ? messages[lastUserIndex].getDate() : null;
        var latestMessage = messages[messages.length - 1];
        var otherPartyReplied = latestMessage.getFrom().toLowerCase().indexOf(userEmail) === -1
                                && lastUserIndex >= 0
                                && lastUserIndex < messages.length - 1;

        if (otherPartyReplied) {
          // The other party replied — remove "Awaiting" so the add-on re-evaluates on next open
          thread.removeLabel(label);
          // Also remove "Needs Follow-up" if present
          var followUpLabel = GmailApp.getUserLabelByName(LABEL_FOLLOWUP);
          if (followUpLabel) thread.removeLabel(followUpLabel);
          resolved++;
          continue;
        }

        // Still awaiting — check if it's overdue
        if (lastUserDate) {
          var daysSince = Math.floor((now - lastUserDate) / (1000 * 60 * 60 * 24));
          if (daysSince >= FOLLOW_UP_AFTER_DAYS) {
            var fuLabel = GmailApp.getUserLabelByName(LABEL_FOLLOWUP);
            if (!fuLabel) {
              var parent = GmailApp.getUserLabelByName(LABEL_PREFIX);
              if (!parent) GmailApp.createLabel(LABEL_PREFIX);
              fuLabel = GmailApp.createLabel(LABEL_FOLLOWUP);
            }
            thread.addLabel(fuLabel);
            nudged++;
          }
        }

      } catch (threadErr) {
        Logger.log('runDailyMeetingScan: error on thread ' + thread.getId() + ' — ' + threadErr.message);
      }
    }

    Logger.log('Daily scan complete. Scanned: ' + scanned + ', Resolved: ' + resolved + ', Nudged: ' + nudged);

  } catch (e) {
    Logger.log('runDailyMeetingScan failed: ' + e.message + '\n' + e.stack);
  }
}
