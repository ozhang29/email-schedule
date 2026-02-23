/**
 * SettingsService.gs
 * Stores and retrieves user settings via PropertiesService.
 * Provides a settings card UI accessible from the homepage.
 */

var SETTINGS_KEY = 'meeting_scheduler_settings';

/**
 * Returns the current settings with defaults applied.
 * @returns {{ autoMode: boolean, userName: string, followUpDays: number }}
 */
function getSettings() {
  var stored = PropertiesService.getUserProperties().getProperty(SETTINGS_KEY);
  var defaults = { autoMode: false, userName: '', followUpDays: 3 };
  if (!stored) return defaults;
  try {
    var parsed = JSON.parse(stored);
    return {
      autoMode: parsed.autoMode === true,
      userName: parsed.userName || '',
      followUpDays: parseInt(parsed.followUpDays, 10) || 3
    };
  } catch (e) {
    return defaults;
  }
}

/**
 * Persists settings to UserProperties.
 * @param {{ autoMode: boolean, userName: string, followUpDays: number }} settings
 */
function saveSettings(settings) {
  PropertiesService.getUserProperties().setProperty(SETTINGS_KEY, JSON.stringify(settings));
}

/**
 * Builds the settings card UI.
 * @returns {Card}
 */
function buildSettingsCard() {
  var settings = getSettings();

  var section = CardService.newCardSection().setHeader('Preferences');

  section.addWidget(CardService.newTextInput()
    .setFieldName('userName')
    .setTitle('Your name (for email sign-offs)')
    .setHint('e.g. Oliver — leave blank to sign off without a name')
    .setValue(settings.userName || ''));

  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName('autoMode')
    .setTitle('Scheduling mode')
    .addItem('Manual — I review and approve each email before it sends', 'false', !settings.autoMode)
    .addItem('Auto — Send emails automatically without my approval', 'true', settings.autoMode));

  section.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('followUpDays')
    .setTitle('Nudge me to follow up after')
    .addItem('2 days', '2', settings.followUpDays === 2)
    .addItem('3 days', '3', settings.followUpDays === 3)
    .addItem('5 days', '5', settings.followUpDays === 5)
    .addItem('7 days', '7', settings.followUpDays === 7));

  section.addWidget(CardService.newTextButton()
    .setText('Save Settings')
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('handleSaveSettings')));

  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle('Meeting Scheduler')
      .setSubtitle('Settings'))
    .addSection(section)
    .build();
}

/**
 * Handles the Save Settings button.
 */
function handleSaveSettings(event) {
  try {
    var fi = (event.commonEventObject || {}).formInputs || {};

    var userName = fi.userName && fi.userName.stringInputs
      ? fi.userName.stringInputs.value[0].trim() : '';

    var autoModeVal = fi.autoMode && fi.autoMode.stringInputs
      ? fi.autoMode.stringInputs.value[0] : 'false';
    var autoMode = autoModeVal === 'true';

    var followUpDaysVal = fi.followUpDays && fi.followUpDays.stringInputs
      ? fi.followUpDays.stringInputs.value[0] : '3';
    var followUpDays = parseInt(followUpDaysVal, 10) || 3;

    var settings = { autoMode: autoMode, userName: userName, followUpDays: followUpDays };
    saveSettings(settings);

    // Install or remove the auto-processor trigger based on mode
    if (autoMode) {
      setupAutoTrigger();
    } else {
      removeAutoTrigger();
    }

    var modeLabel = autoMode ? 'Auto' : 'Manual';
    var nameNote = userName ? ', signing as <b>' + userName + '</b>' : '';
    var confirmText = 'Settings saved! Mode: <b>' + modeLabel + '</b>' + nameNote + '.';
    if (autoMode) {
      confirmText += '<br><br>Auto mode is on. The assistant will now check for scheduling emails every 10 minutes and send replies automatically. Run <b>setupAutoTrigger()</b> from the Apps Script editor if the background trigger is not yet installed.';
    }

    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .pushCard(CardService.newCardBuilder()
          .setHeader(CardService.newCardHeader()
            .setTitle('Meeting Scheduler')
            .setSubtitle('Settings Saved'))
          .addSection(CardService.newCardSection()
            .addWidget(CardService.newTextParagraph().setText(confirmText)))
          .build()))
      .build();

  } catch (e) {
    Logger.log('handleSaveSettings error: ' + e.message + '\n' + e.stack);
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().pushCard(buildErrorCard(e.message)))
      .build();
  }
}
