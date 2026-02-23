/**
 * Code.gs
 * Entry point, CardService UI, and action handlers for the Meeting Scheduler add-on.
 */

/**
 * Gmail contextual trigger — builds the add-on card for the open email thread.
 * @param {Object} event - The Gmail add-on event object.
 * @returns {Card[]}
 */
function buildAddOn(event) {
  var messageId = event.gmail && event.gmail.messageId;
  if (!messageId) {
    return [buildErrorCard('Could not read the current email.')];
  }

  try {
    var threadContent = getThreadContent(messageId);
    var analysis = analyzeMeetingThread(threadContent);

    if (!analysis.isSchedulingThread) {
      return [buildNotSchedulingCard()];
    }

    var inviteCheck = checkExistingInvite(analysis);
    if (inviteCheck.alreadyScheduled || analysis.meetingStatus === 'already_scheduled') {
      return [buildAlreadyScheduledCard(analysis, inviteCheck)];
    }

    if (analysis.meetingStatus === 'inbound_request') {
      // If sender proposed times with parseable ISO dates, check calendar availability
      var hasParseable = (analysis.proposedTimes || []).some(function(pt) {
        return pt.startIso && /^\d{4}-\d{2}-\d{2}T/.test(pt.startIso);
      });

      if (hasParseable) {
        var proposedAvailability = checkProposedTimes(analysis.proposedTimes, analysis.durationMinutes || 30);
        // Cache large objects to stay under 2KB action-parameter limit
        var cache = CacheService.getUserCache();
        var cacheKey = Utilities.getUuid();
        cache.put(cacheKey, JSON.stringify({ analysis: analysis, proposedAvailability: proposedAvailability, messageId: messageId }), 600);
        return [buildInboundTimesCard(analysis, proposedAvailability, messageId, cacheKey)];
      }

      return [buildInboundRequestCard(analysis, messageId)];
    }

    if (analysis.meetingStatus === 'user_promised_times') {
      return [buildSendAvailabilityCard(analysis, messageId)];
    }

    if (analysis.meetingStatus === 'agreement_reached') {
      return [buildReadyToSendCard(analysis, messageId)];
    }

    // Default: no_agreement
    return [buildNoAgreementCard(analysis)];

  } catch (e) {
    Logger.log('buildAddOn error: ' + e.message + '\n' + e.stack);
    return [buildErrorCard(e.message)];
  }
}

/**
 * Homepage trigger shown outside a specific email context.
 */
function buildHomepage() {
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Meeting Scheduler')
      .setSubtitle('Powered by AI'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('Open any scheduling email thread and this add-on will automatically detect meeting details.\n\n' +
                 'When both parties have agreed on a time, click <b>Send Google Meet Invite</b> to create a calendar event and send invitations instantly.'))
      .addWidget(CardService.newTextParagraph()
        .setText('<b>How it works:</b>\n' +
                 '1. Open a scheduling email in Gmail\n' +
                 '2. The add-on reads the thread\n' +
                 '3. If an agreement is found, send the invite with one click')));

  return card.build();
}

// ---------------------------------------------------------------------------
// Card builders — informational (no form)
// ---------------------------------------------------------------------------

function buildNotSchedulingCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('This email doesn\'t appear to involve meeting coordination.')))
    .build();
}

function buildNoAgreementCard(analysis) {
  var section = CardService.newCardSection()
    .setHeader('No Agreement Yet')
    .addWidget(CardService.newTextParagraph()
      .setText(analysis.summary || 'Scheduling is in progress but no time has been agreed upon yet.'));

  if (analysis.proposedTimes && analysis.proposedTimes.length > 0) {
    section.addWidget(CardService.newTextParagraph().setText('<b>Proposed times:</b>'));
    for (var i = 0; i < analysis.proposedTimes.length; i++) {
      var pt = analysis.proposedTimes[i];
      section.addWidget(CardService.newKeyValue()
        .setTopLabel(pt.proposedBy || 'Proposed')
        .setContent(pt.displayText || pt.startIso || 'Unknown time'));
    }
  }

  if (analysis.participantEmails && analysis.participantEmails.length > 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<b>Participants:</b> ' + analysis.participantEmails.join(', ')));
  }

  if (analysis.uncertainty) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<i>Note: ' + analysis.uncertainty + '</i>'));
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

function buildAlreadyScheduledCard(analysis, inviteCheck) {
  var section = CardService.newCardSection()
    .setHeader('Already Scheduled')
    .addWidget(CardService.newTextParagraph()
      .setText(analysis.summary || 'A calendar invite has already been sent for this meeting.'));

  if (analysis.agreedTime && analysis.agreedTime.displayText) {
    section.addWidget(CardService.newKeyValue()
      .setTopLabel('Confirmed time')
      .setContent(analysis.agreedTime.displayText));
  }

  if (analysis.participantEmails && analysis.participantEmails.length > 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<b>Participants:</b> ' + analysis.participantEmails.join(', ')));
  }

  var eventUrl = inviteCheck.eventUrl;
  if (eventUrl) {
    section.addWidget(CardService.newTextButton()
      .setText('View in Calendar')
      .setOpenLink(CardService.newOpenLink().setUrl(eventUrl)));
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

function buildReadyToSendCard(analysis, messageId) {
  var section = CardService.newCardSection()
    .setHeader('Agreement Reached')
    .addWidget(CardService.newTextParagraph()
      .setText(analysis.summary || 'Both parties have agreed on a meeting time.'));

  if (analysis.agreedTime && analysis.agreedTime.displayText) {
    section.addWidget(CardService.newKeyValue()
      .setTopLabel('Agreed time')
      .setContent(analysis.agreedTime.displayText));
  }

  if (analysis.participantEmails && analysis.participantEmails.length > 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<b>Participants:</b> ' + analysis.participantEmails.join(', ')));
  }

  if (analysis.uncertainty) {
    section.addWidget(CardService.newTextParagraph()
      .setText('<i>Note: ' + analysis.uncertainty + '</i>'));
  }

  // Pass analysis as JSON so handleSendInvite doesn't need a second AI call
  var analysisJson = JSON.stringify(analysis);

  section.addWidget(CardService.newTextButton()
    .setText('Send Google Meet Invite')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleSendInvite')
      .setParameters({ analysisJson: analysisJson, messageId: messageId })));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

function buildSuccessCard(createdEvent) {
  var section = CardService.newCardSection()
    .setHeader('Invite Sent!')
    .addWidget(CardService.newTextParagraph()
      .setText('Calendar invitations have been sent to all participants.'));

  if (createdEvent.displayTime) {
    section.addWidget(CardService.newKeyValue()
      .setTopLabel('Time')
      .setContent(createdEvent.displayTime));
  }

  if (createdEvent.meetUrl) {
    section.addWidget(CardService.newTextButton()
      .setText('Join Google Meet')
      .setOpenLink(CardService.newOpenLink().setUrl(createdEvent.meetUrl)));
  }

  if (createdEvent.eventUrl) {
    section.addWidget(CardService.newTextButton()
      .setText('View in Calendar')
      .setOpenLink(CardService.newOpenLink().setUrl(createdEvent.eventUrl)));
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

function buildErrorCard(msg) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText('<b>Error:</b> ' + (msg || 'An unexpected error occurred. Check Stackdriver logs.'))))
    .build();
}

// ---------------------------------------------------------------------------
// Card builders — availability forms
// ---------------------------------------------------------------------------

/**
 * Shared helper: builds the duration dropdown + manual availability form used by
 * buildInboundRequestCard and buildSendAvailabilityCard.
 * Uses CacheService to store analysis (avoids 2KB action-parameter limit).
 */
function buildAvailabilityCard(analysis, messageId, sectionHeader, summaryText, buttonText) {
  var section = CardService.newCardSection()
    .setHeader(sectionHeader)
    .addWidget(CardService.newTextParagraph().setText(summaryText));

  if (analysis.requesterName) {
    section.addWidget(CardService.newKeyValue()
      .setTopLabel('From')
      .setContent(analysis.requesterName));
  }

  if (analysis.proposedTimes && analysis.proposedTimes.length > 0) {
    section.addWidget(CardService.newTextParagraph().setText('<b>Their proposed times:</b>'));
    for (var i = 0; i < analysis.proposedTimes.length; i++) {
      var pt = analysis.proposedTimes[i];
      section.addWidget(CardService.newKeyValue()
        .setTopLabel(pt.proposedBy || 'Proposed')
        .setContent(pt.displayText || pt.startIso || 'Unknown time'));
    }
  }

  // Duration dropdown — 30 min selected by default
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('duration')
    .setTitle('Duration')
    .addItem('30 minutes', '30', true)
    .addItem('60 minutes', '60', false)
    .addItem('90 minutes', '90', false));

  // Manual availability text input
  section.addWidget(CardService.newTextInput()
    .setFieldName('manualAvailability')
    .setTitle('Your availability (optional)')
    .setHint('e.g. Mon 2–4 pm, Tue after noon — leave blank to use calendar')
    .setMultiline(true));

  // Store analysis in cache to avoid 2KB parameter limit
  var cache = CacheService.getUserCache();
  var cacheKey = Utilities.getUuid();
  cache.put(cacheKey, JSON.stringify({ analysis: analysis, messageId: messageId }), 600);

  section.addWidget(CardService.newTextButton()
    .setText(buttonText)
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleReplyWithAvailability')
      .setParameters({ cacheKey: cacheKey })));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

function buildInboundRequestCard(analysis, messageId) {
  return buildAvailabilityCard(
    analysis, messageId,
    'Meeting Request',
    analysis.summary || 'Someone has asked you to find a time to meet.',
    'Preview Reply'
  );
}

function buildSendAvailabilityCard(analysis, messageId) {
  return buildAvailabilityCard(
    analysis, messageId,
    'Follow Up with Availability',
    analysis.summary || 'You promised to send your availability. Share your times now.',
    'Preview Reply'
  );
}

/**
 * Shows sender's proposed times with ✓ Free / ✗ Busy badges checked against the user's calendar.
 * User types a natural language response ("ok", "the first one works", etc.).
 */
function buildInboundTimesCard(analysis, proposedAvailability, messageId, cacheKey) {
  var section = CardService.newCardSection()
    .setHeader('Meeting Request')
    .addWidget(CardService.newTextParagraph()
      .setText(analysis.summary || 'Check which of their proposed times works for you:'));

  if (analysis.requesterName) {
    section.addWidget(CardService.newKeyValue()
      .setTopLabel('From')
      .setContent(analysis.requesterName));
  }

  section.addWidget(CardService.newTextParagraph().setText('<b>Their proposed times:</b>'));
  for (var i = 0; i < proposedAvailability.length; i++) {
    var t = proposedAvailability[i];
    var label = t.availLabel || ('Option ' + (i + 1));
    section.addWidget(CardService.newKeyValue()
      .setTopLabel(label)
      .setContent(t.displayText || 'Unknown time'));
  }

  // Natural language input — pre-filled with "ok" for one-tap acceptance
  section.addWidget(CardService.newTextInput()
    .setFieldName('naturalResponse')
    .setTitle('Your response')
    .setHint('e.g. "ok", "the first one works", "let\'s do option 2"')
    .setValue('ok'));

  // Duration dropdown (needed when a calendar invite is created)
  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('duration')
    .setTitle('Meeting duration')
    .addItem('30 minutes', '30', true)
    .addItem('60 minutes', '60', false)
    .addItem('90 minutes', '90', false));

  section.addWidget(CardService.newTextButton()
    .setText('Preview Response')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleNaturalResponse')
      .setParameters({ cacheKey: cacheKey })));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

/**
 * Shows the AI-drafted email in an editable text field. User can edit then send.
 * @param {string} draftBody - AI-generated email text.
 * @param {string} draftCacheKey - Cache key where send context (mode, messageId, etc.) is stored.
 */
function buildDraftPreviewCard(draftBody, draftCacheKey) {
  var section = CardService.newCardSection()
    .setHeader('Preview Draft')
    .addWidget(CardService.newTextParagraph()
      .setText('Review and edit before sending:'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('editedDraft')
    .setTitle('Email')
    .setValue(draftBody)
    .setMultiline(true));

  section.addWidget(CardService.newTextButton()
    .setText('Send')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleSendDraft')
      .setParameters({ draftCacheKey: draftCacheKey })));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

/**
 * Confirmation card shown after reply is sent.
 * @param {Array} freeSlots - Proposed slots (availability mode).
 * @param {string} manualAvailability - Manual text (availability mode).
 * @param {Object|null} createdEvent - Calendar event info (acceptance mode), or null.
 */
function buildReplySentCard(freeSlots, manualAvailability, createdEvent) {
  var section = CardService.newCardSection()
    .setHeader('Reply Sent!');

  if (createdEvent) {
    section.addWidget(CardService.newTextParagraph()
      .setText('Your acceptance has been sent and a Google Calendar invite has been created.'));

    if (createdEvent.displayTime) {
      section.addWidget(CardService.newKeyValue()
        .setTopLabel('Time')
        .setContent(createdEvent.displayTime));
    }

    if (createdEvent.meetUrl) {
      section.addWidget(CardService.newTextButton()
        .setText('Join Google Meet')
        .setOpenLink(CardService.newOpenLink().setUrl(createdEvent.meetUrl)));
    }

    if (createdEvent.eventUrl) {
      section.addWidget(CardService.newTextButton()
        .setText('View in Calendar')
        .setOpenLink(CardService.newOpenLink().setUrl(createdEvent.eventUrl)));
    }

  } else if (manualAvailability) {
    section.addWidget(CardService.newTextParagraph()
      .setText('Your availability has been sent. You shared:'));
    section.addWidget(CardService.newTextParagraph().setText(manualAvailability));

  } else if (freeSlots && freeSlots.length > 0) {
    section.addWidget(CardService.newTextParagraph()
      .setText('Your availability has been sent. You proposed:'));
    for (var i = 0; i < freeSlots.length; i++) {
      section.addWidget(CardService.newKeyValue()
        .setTopLabel('Option ' + (i + 1))
        .setContent(freeSlots[i].displayText));
    }

  } else {
    section.addWidget(CardService.newTextParagraph().setText('Your reply has been sent.'));
  }

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

// ---------------------------------------------------------------------------
// Action handlers
// ---------------------------------------------------------------------------

/**
 * Handles "Preview Response" on the inbound times card.
 * Interprets natural language, drafts an acceptance email, shows preview.
 */
function handleNaturalResponse(event) {
  try {
    var params = event.commonEventObject ? event.commonEventObject.parameters : event.parameters;
    var cacheKey = params.cacheKey;
    if (!cacheKey) throw new Error('Missing cacheKey. Please reopen the add-on.');

    var cache = CacheService.getUserCache();
    var cached = cache.get(cacheKey);
    if (!cached) throw new Error('Session expired. Please reopen the add-on.');

    var stored = JSON.parse(cached);
    var analysis = stored.analysis;
    var proposedAvailability = stored.proposedAvailability;
    var messageId = stored.messageId;

    var fi = (event.commonEventObject || {}).formInputs || {};
    var naturalText = fi.naturalResponse && fi.naturalResponse.stringInputs
      ? fi.naturalResponse.stringInputs.value[0] : 'ok';
    var durationStr = fi.duration && fi.duration.stringInputs
      ? fi.duration.stringInputs.value[0] : '30';
    var durationMinutes = parseInt(durationStr, 10) || 30;

    // Interpret which time the user chose
    var interpretation = interpretNaturalResponse(naturalText, proposedAvailability);
    var chosenTime = interpretation.chosenTime;

    // Draft the acceptance email (reads thread for real names)
    var threadContent = getThreadContent(messageId);
    var draftBody = draftAcceptanceEmail(threadContent, chosenTime, analysis);

    // Store context for handleSendDraft
    var draftCacheKey = Utilities.getUuid();
    cache.put(draftCacheKey, JSON.stringify({
      mode: 'acceptance',
      messageId: messageId,
      chosenTime: chosenTime,
      durationMinutes: durationMinutes,
      meetingTitle: analysis.meetingTitle || 'Meeting',
      participantEmails: analysis.participantEmails || []
    }), 600);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildDraftPreviewCard(draftBody, draftCacheKey)))
      .build();

  } catch (e) {
    Logger.log('handleNaturalResponse error: ' + e.message + '\n' + e.stack);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildErrorCard(e.message)))
      .build();
  }
}

/**
 * Handles the availability form button click.
 * Gets free slots (or uses manual text), drafts email, shows preview — does NOT send yet.
 */
function handleReplyWithAvailability(event) {
  try {
    var params = event.commonEventObject ? event.commonEventObject.parameters : event.parameters;

    var cache = CacheService.getUserCache();
    var analysis, messageId;

    if (params.cacheKey) {
      var cached = cache.get(params.cacheKey);
      if (!cached) throw new Error('Session expired. Please reopen the add-on.');
      var stored = JSON.parse(cached);
      analysis = stored.analysis;
      messageId = stored.messageId;
    } else {
      // Legacy fallback: analysisJson passed directly
      if (!params.analysisJson || !params.messageId) {
        throw new Error('Missing required parameters. Please refresh the add-on.');
      }
      analysis = JSON.parse(params.analysisJson);
      messageId = params.messageId;
    }

    var fi = (event.commonEventObject || {}).formInputs || {};
    var durationStr = fi.duration && fi.duration.stringInputs
      ? fi.duration.stringInputs.value[0] : '30';
    var durationMinutes = parseInt(durationStr, 10) || 30;

    // Override AI-detected duration with the user's form selection
    analysis.durationMinutes = durationMinutes;

    var manualAvailability = fi.manualAvailability && fi.manualAvailability.stringInputs
      ? fi.manualAvailability.stringInputs.value[0] : '';

    var freeSlots = [];
    if (!manualAvailability) {
      freeSlots = getFreeSlots(durationMinutes, 3);
      if (freeSlots.length === 0) {
        throw new Error('No free slots found in the next 14 days between 9am–5pm.');
      }
    }

    var threadContent = getThreadContent(messageId);
    var draftBody = draftReplyEmail(threadContent, freeSlots, analysis, manualAvailability, durationMinutes);

    // Store context for handleSendDraft
    var draftCacheKey = Utilities.getUuid();
    cache.put(draftCacheKey, JSON.stringify({
      mode: 'availability',
      messageId: messageId,
      freeSlots: freeSlots,
      manualAvailability: manualAvailability
    }), 600);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildDraftPreviewCard(draftBody, draftCacheKey)))
      .build();

  } catch (e) {
    Logger.log('handleReplyWithAvailability error: ' + e.message + '\n' + e.stack);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildErrorCard(e.message)))
      .build();
  }
}

/**
 * Handles the "Send" button on the draft preview card.
 * Sends the (possibly edited) email. For acceptance mode, also creates a calendar invite.
 */
function handleSendDraft(event) {
  try {
    var params = event.commonEventObject ? event.commonEventObject.parameters : event.parameters;
    var draftCacheKey = params.draftCacheKey;
    if (!draftCacheKey) throw new Error('Missing draftCacheKey. Please reopen the add-on.');

    var cache = CacheService.getUserCache();
    var cached = cache.get(draftCacheKey);
    if (!cached) throw new Error('Session expired. Please reopen the add-on.');

    var stored = JSON.parse(cached);
    var mode = stored.mode;
    var messageId = stored.messageId;

    var fi = (event.commonEventObject || {}).formInputs || {};
    var emailBody = fi.editedDraft && fi.editedDraft.stringInputs
      ? fi.editedDraft.stringInputs.value[0] : '';

    if (!emailBody || !emailBody.trim()) {
      throw new Error('Email body is empty. Please type a message before sending.');
    }

    sendReply(messageId, emailBody.trim());

    if (mode === 'acceptance') {
      var chosenTime = stored.chosenTime;
      var durationMinutes = stored.durationMinutes || 30;
      var inviteAnalysis = {
        agreedTime: {
          startIso: chosenTime.startIso,
          endIso: chosenTime.endIso || new Date(new Date(chosenTime.startIso).getTime() + durationMinutes * 60000).toISOString(),
          displayText: chosenTime.displayText,
          timezone: 'America/New_York'
        },
        durationMinutes: durationMinutes,
        meetingTitle: stored.meetingTitle || 'Meeting',
        participantEmails: stored.participantEmails || []
      };

      var createdEvent = null;
      try {
        createdEvent = createMeetEvent(inviteAnalysis);
      } catch (inviteErr) {
        Logger.log('Calendar invite creation failed (non-fatal): ' + inviteErr.message);
      }

      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(buildReplySentCard([], '', createdEvent)))
        .build();

    } else {
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(buildReplySentCard(stored.freeSlots || [], stored.manualAvailability || '', null)))
        .build();
    }

  } catch (e) {
    Logger.log('handleSendDraft error: ' + e.message + '\n' + e.stack);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildErrorCard(e.message)))
      .build();
  }
}

/**
 * Handles the "Send Google Meet Invite" button click (agreement_reached flow).
 */
function handleSendInvite(event) {
  try {
    var params = event.commonEventObject ? event.commonEventObject.parameters : event.parameters;
    var analysisJson = params.analysisJson;
    if (!analysisJson) {
      throw new Error('Missing analysis data. Please refresh the add-on.');
    }

    var analysis = JSON.parse(analysisJson);
    var createdEvent = createMeetEvent(analysis);

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildSuccessCard(createdEvent)))
      .build();

  } catch (e) {
    Logger.log('handleSendInvite error: ' + e.message + '\n' + e.stack);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildErrorCard(e.message)))
      .build();
  }
}
