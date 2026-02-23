/**
 * ComposeService.gs
 * Gmail compose trigger — lets the user generate and insert a scheduling
 * email (with available times) directly into a new compose window.
 *
 * Manifest entry required in appsscript.json → addOns → gmail:
 *   "composeTrigger": {
 *     "selectActions": [{ "text": "Insert my availability", "runFunction": "buildComposeCard" }],
 *     "draftAccess": "METADATA"
 *   }
 */

/**
 * Compose trigger entry point. Returns the sidebar card shown when the user
 * opens the add-on from within a compose window.
 * @param {Object} event - Gmail compose event.
 * @returns {Card}
 */
function buildComposeCard(event) {
  var section = CardService.newCardSection()
    .setHeader('Generate Scheduling Email')
    .addWidget(CardService.newTextParagraph()
      .setText('Add optional context and your available times will be inserted automatically.'));

  section.addWidget(CardService.newTextInput()
    .setFieldName('contextMessage')
    .setTitle('Opening message (optional)')
    .setHint('e.g. "Great to meet you today — I\'d love to chat more."')
    .setMultiline(true));

  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('duration')
    .setTitle('Meeting duration')
    .addItem('30 minutes', '30', true)
    .addItem('60 minutes', '60', false)
    .addItem('90 minutes', '90', false));

  section.addWidget(CardService.newTextButton()
    .setText('Generate & Insert')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleInsertAvailability')));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Meeting Scheduler'))
    .addSection(section)
    .build();
}

/**
 * Handles "Generate & Insert": fetches free calendar slots, drafts the full
 * outgoing email, and inserts it into the compose window.
 * @param {Object} event - The action event from the card button.
 * @returns {ActionResponse} with UpdateDraftActionResponse.
 */
function handleInsertAvailability(event) {
  try {
    var settings = getSettings();
    var fi = (event.commonEventObject || {}).formInputs || {};

    var contextMessage = fi.contextMessage && fi.contextMessage.stringInputs
      ? fi.contextMessage.stringInputs.value[0].trim() : '';

    var durationMinutes = parseInt(
      (fi.duration && fi.duration.stringInputs ? fi.duration.stringInputs.value[0] : '30'),
      10
    ) || 30;

    var freeSlots = getFreeSlots(durationMinutes, 3);
    if (freeSlots.length === 0) {
      throw new Error('No free slots found in the next 14 days between 9am–5pm. Check your calendar.');
    }

    var emailBody = draftOutboundEmail(
      contextMessage,
      freeSlots,
      durationMinutes,
      settings.userName || ''
    );

    // Convert plain text to simple HTML so newlines render correctly in Gmail
    var htmlBody = emailBody
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\n/g, '<br>');

    var updateDraftBodyAction = CardService.newUpdateDraftBodyAction()
      .addUpdateContent(htmlBody, CardService.ContentType.MUTABLE_HTML)
      .setUpdateType(CardService.UpdateDraftBodyType.IN_PLACE_INSERT);

    return CardService.newActionResponseBuilder()
      .setUpdateDraftActionResponse(
        CardService.newUpdateDraftActionResponseBuilder()
          .setUpdateDraftBodyAction(updateDraftBodyAction)
          .build()
      )
      .build();

  } catch (e) {
    Logger.log('handleInsertAvailability error: ' + e.message + '\n' + e.stack);
    // In compose context, navigation to an error card is the best we can do
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(buildErrorCard(e.message)))
      .build();
  }
}
