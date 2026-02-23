/**
 * ClaudeService.gs
 * Calls OpenAI API to analyze scheduling email threads.
 */

var AI_API_URL = 'https://api.openai.com/v1/chat/completions';
var AI_MODEL = 'gpt-4o';
var AI_MAX_TOKENS = 1024;

/**
 * Analyzes a meeting thread and returns a MeetingAnalysis object.
 * @param {{ transcript: string, participants: Array<{name:string,email:string}>, hasIcsAttachment: boolean, subject: string }} threadContent
 * @returns {MeetingAnalysis}
 */
function analyzeMeetingThread(threadContent) {
  var today = new Date().toISOString().split('T')[0];

  var systemPrompt =
    'You are a meeting scheduling assistant. Analyze the email thread and return ONLY a valid JSON object ' +
    'with no markdown, no code fences, no extra text. Today\'s date is ' + today + '. ' +
    'Use this to resolve relative dates like "this Friday" or "tomorrow".\n\n' +
    'Return this exact JSON structure:\n' +
    '{\n' +
    '  "isSchedulingThread": boolean,\n' +
    '  "meetingStatus": "inbound_request" | "no_agreement" | "agreement_reached" | "already_scheduled" | "user_promised_times" | "awaiting_response",\n' +
    '  "requesterName": "name of person asking to meet" | null,\n' +
    '  "proposedTimes": [{ "proposedBy": "email", "displayText": "...", "startIso": "ISO8601 or null" }],\n' +
    '  "agreedTime": { "startIso": "ISO8601", "endIso": "ISO8601", "timezone": "IANA timezone", "displayText": "human-readable" } | null,\n' +
    '  "participantEmails": ["email1", "email2"],\n' +
    '  "calendarInviteSent": boolean,\n' +
    '  "meetingTitle": "short descriptive title",\n' +
    '  "durationMinutes": 30,\n' +
    '  "summary": "1-2 sentence summary of the scheduling situation",\n' +
    '  "uncertainty": "note about ambiguity if any" | null\n' +
    '}\n\n' +
    'Rules:\n' +
    '- Set meetingStatus to "inbound_request" when the most recent email is someone asking the user to find a time to meet and the user has not yet proposed specific times back.\n' +
    '- Set meetingStatus to "awaiting_response" when the user\'s most recent email in the thread proposes specific times or shares availability, AND that email is the last message in the thread (the other party has not replied yet).\n' +
    '- Set meetingStatus to "user_promised_times" when the user\'s most recent message promises to send availability (e.g. "I\'ll follow up with some times", "let me find some slots", "I\'ll send a few options") but has not yet done so.\n' +
    '- Set meetingStatus to "agreement_reached" ONLY when both parties have explicitly confirmed the SAME specific time.\n' +
    '- Set meetingStatus to "already_scheduled" if you see language like "I\'ve sent a calendar invite", "check your calendar", "calendar invite sent", or an .ics attachment was noted.\n' +
    '- Set calendarInviteSent to true if invite language is present or an .ics attachment was noted.\n' +
    '- If isSchedulingThread is false, other fields may be empty/null.\n' +
    '- For agreedTime, always include both startIso and endIso. Use durationMinutes to compute endIso if not explicit.\n' +
    '- Prefer the most recent timezone mentioned; default to America/New_York if unknown.\n' +
    '- Return ONLY the JSON object.';

  var userPrompt =
    'Email thread subject: ' + threadContent.subject + '\n' +
    (threadContent.hasIcsAttachment ? 'Note: This thread has an .ics calendar attachment.\n' : '') +
    '\n--- EMAIL THREAD ---\n' +
    threadContent.transcript;

  var rawResponse = callAiApi(systemPrompt, userPrompt);

  try {
    var analysis = JSON.parse(rawResponse);
    // Attach hasIcsAttachment for CalendarService
    analysis.hasIcsAttachment = threadContent.hasIcsAttachment;
    return analysis;
  } catch (e) {
    // Attempt to extract JSON from response if there's surrounding text
    var jsonMatch = rawResponse.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      var analysis = JSON.parse(jsonMatch[0]);
      analysis.hasIcsAttachment = threadContent.hasIcsAttachment;
      return analysis;
    }
    throw new Error('AI returned invalid JSON: ' + rawResponse.substring(0, 200));
  }
}

/**
 * Drafts a brief reply email proposing free slots or manual availability text.
 * @param {{ transcript: string, subject: string }} threadContent
 * @param {Array<{ startIso: string, endIso: string, displayText: string }>} freeSlots
 * @param {MeetingAnalysis} analysis
 * @param {string} [manualAvailability] - If non-empty, use this text instead of freeSlots.
 * @param {number} [durationMinutes] - Form-selected duration; overrides analysis value.
 * @returns {string} Plain-text email body.
 */
function draftReplyEmail(threadContent, freeSlots, analysis, manualAvailability, durationMinutes) {
  var effectiveDuration = durationMinutes || analysis.durationMinutes || 30;

  var systemPrompt =
    'You are drafting a professional reply on behalf of the user. ' +
    'Format:\n' +
    '1. One short opening line — thank them for reaching out and say you are sharing some available windows.\n' +
    '2. The availability as a bulleted list (use a dash "-" for each bullet). Each bullet is one availability window exactly as provided — do not rephrase or merge them.\n' +
    '3. One short closing line — ask them to pick a time that works.\n' +
    '4. Sign-off and name (see rules below).\n' +
    'Rules for names:\n' +
    '(1) To address the recipient: look in the thread for their first name from their own signature or sign-off. Use it in the opening if clearly found. Otherwise just write "Hi" with no name.\n' +
    '(2) To sign off: look in the thread for how the user themselves signs their emails. Use that sign-off if found. If not found, end with just "Best," or "Thanks," with no name.\n' +
    '(3) NEVER write placeholder text like [Your Name], [Name], or anything in square brackets.\n' +
    'Return only the email body text, no subject line.';

  var transcript = threadContent.transcript.substring(0, 3000);
  var userPrompt;

  if (manualAvailability) {
    // Format manual text as bullet list if not already bulleted
    var manualLines = manualAvailability.split(/[\n,;]+/).map(function(s) { return s.trim(); }).filter(Boolean);
    var manualBullets = manualLines.map(function(l) { return '- ' + l; }).join('\n');

    userPrompt =
      'Email subject: ' + threadContent.subject + '\n' +
      'Duration: ' + effectiveDuration + ' minutes\n\n' +
      'Thread (find real names from signatures here):\n' + transcript + '\n\n' +
      'Availability windows to list as bullets:\n' + manualBullets + '\n\n' +
      'Write the reply using the format described. List each window as its own bullet.';
  } else {
    var slotBullets = freeSlots.map(function(s) {
      return '- ' + s.displayText;
    }).join('\n');

    userPrompt =
      'Email subject: ' + threadContent.subject + '\n' +
      'Duration: ' + effectiveDuration + ' minutes\n\n' +
      'Thread (find real names from signatures here):\n' + transcript + '\n\n' +
      'Availability windows to list as bullets:\n' + slotBullets + '\n\n' +
      'Write the reply using the format described. List each window as its own bullet. Do not rephrase the times.';
  }

  return callAiApi(systemPrompt, userPrompt);
}

/**
 * Drafts a short acceptance email confirming a specific proposed time.
 * @param {{ transcript: string, subject: string }} threadContent
 * @param {{ displayText: string, startIso: string }} chosenTime
 * @param {MeetingAnalysis} analysis
 * @returns {string} Plain-text email body.
 */
function draftAcceptanceEmail(threadContent, chosenTime, analysis) {
  var systemPrompt =
    'You are drafting a short professional acceptance email on behalf of the user. ' +
    'Keep it to 2-3 sentences. ' +
    'Rules for names: ' +
    '(1) To address the recipient: look in the thread for their first name from their own signature or sign-off. Use it if clearly found. Otherwise just write "Hi" with no name. ' +
    '(2) To sign off: look in the thread for how the user themselves signs their emails. Use that if found. If not, end with just "Best," or "Thanks," with no name at all. ' +
    '(3) NEVER write placeholder text like [Your Name], [Name], or anything in square brackets. ' +
    'Return only the email body text, no subject line.';

  var userPrompt =
    'Email subject: ' + threadContent.subject + '\n\n' +
    'Thread (find real names from signatures here):\n' + threadContent.transcript.substring(0, 3000) + '\n\n' +
    'The user is accepting this time: ' + chosenTime.displayText + '\n\n' +
    'Write a short confirmation reply saying this time works. Mention the specific time. 2-3 sentences.';

  return callAiApi(systemPrompt, userPrompt);
}

/**
 * Interprets natural language to pick which proposed time the user wants to accept.
 * Fast path: generic affirmatives ("ok", "sure", etc.) pick the first free slot without an AI call.
 * @param {string} userText
 * @param {Array<{ displayText:string, startIso:string, free:boolean|null, availLabel:string }>} proposedAvailability
 * @returns {{ chosenIndex: number, chosenTime: object }}
 */
function interpretNaturalResponse(userText, proposedAvailability) {
  // Fast path: generic affirmative → first free slot, no AI call needed
  var genericOk = /^(ok|okay|good|great|sure|works|fine|sounds good|perfect|yes|yep|yeah|that works|go ahead|do it|either works|any works)[\.\!]?$/i;
  if (!userText || genericOk.test(userText.trim())) {
    for (var i = 0; i < proposedAvailability.length; i++) {
      if (proposedAvailability[i].free === true) {
        return { chosenIndex: i, chosenTime: proposedAvailability[i] };
      }
    }
    return { chosenIndex: 0, chosenTime: proposedAvailability[0] };
  }

  // AI path: specific references like "the first one", "pick Monday", "let's do option 2"
  var slotDescriptions = proposedAvailability.map(function(pt, i) {
    var status = pt.availLabel ? ' (' + pt.availLabel + ')' : '';
    return (i + 1) + '. ' + pt.displayText + status;
  }).join('\n');

  var systemPrompt =
    'You pick which proposed meeting time the user is referring to. ' +
    'Return ONLY valid JSON: {"chosenIndex": N} where N is the 0-based index. No markdown, no extra text.';

  var userPrompt =
    'User said: "' + userText + '"\n\n' +
    'Proposed times:\n' + slotDescriptions + '\n\n' +
    'Return the 0-based index. If unclear, return 0.';

  try {
    var raw = callAiApi(systemPrompt, userPrompt);
    var match = raw.match(/\{[\s\S]*?\}/);
    if (match) {
      var result = JSON.parse(match[0]);
      var idx = parseInt(result.chosenIndex, 10);
      if (!isNaN(idx) && idx >= 0 && idx < proposedAvailability.length) {
        return { chosenIndex: idx, chosenTime: proposedAvailability[idx] };
      }
    }
  } catch (e) {
    Logger.log('interpretNaturalResponse AI error: ' + e.message);
  }

  // Fallback: first free slot
  for (var i = 0; i < proposedAvailability.length; i++) {
    if (proposedAvailability[i].free === true) {
      return { chosenIndex: i, chosenTime: proposedAvailability[i] };
    }
  }
  return { chosenIndex: 0, chosenTime: proposedAvailability[0] };
}

/**
 * Drafts a polite follow-up email when awaiting a response.
 * @param {{ transcript: string, subject: string }} threadContent
 * @param {number} daysSince - Days since the user's last email.
 * @returns {string} Plain-text email body.
 */
function draftFollowUpEmail(threadContent, daysSince) {
  var systemPrompt =
    'You are drafting a short, polite follow-up email on behalf of the user. ' +
    'Keep it to 2-3 sentences. Friendly and professional — not pushy. ' +
    'Rules for names: ' +
    '(1) Address the recipient by their first name if clearly found in their thread signatures. Otherwise write "Hi" with no name. ' +
    '(2) Sign off with the name the user signs with in the thread, if found. Otherwise just "Best," or "Thanks," with no name. ' +
    '(3) NEVER write placeholder text like [Your Name], [Name], or anything in square brackets. ' +
    'Return only the email body, no subject line.';

  var dayDesc = daysSince <= 1 ? 'a day or two' : daysSince + ' days';
  var userPrompt =
    'Thread (find real names from signatures here):\n' + threadContent.transcript.substring(0, 3000) + '\n\n' +
    'The user sent their availability ' + dayDesc + ' ago and has not received a response. ' +
    'Draft a short, friendly follow-up just checking in. 2-3 sentences.';

  return callAiApi(systemPrompt, userPrompt);
}

/**
 * Makes a raw request to the OpenAI Chat Completions API.
 * @param {string} systemPrompt
 * @param {string} userPrompt
 * @returns {string} The text content of the model's response.
 */
function callAiApi(systemPrompt, userPrompt) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY not set in Script Properties.');
  }

  var payload = {
    model: AI_MODEL,
    max_tokens: AI_MAX_TOKENS,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(AI_API_URL, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('OpenAI API error ' + responseCode + ': ' + responseText.substring(0, 300));
  }

  var responseJson = JSON.parse(responseText);
  if (!responseJson.choices || !responseJson.choices[0] || !responseJson.choices[0].message) {
    throw new Error('Unexpected OpenAI response structure: ' + responseText.substring(0, 200));
  }

  return responseJson.choices[0].message.content.trim();
}
