/**
 * Google Meet Attendance Automation
 *
 * Deployment configuration is stored in Apps Script
 * Project Settings -> Script Properties.
 *
 * Required properties:
 *   RECURRING_EVENT_ID
 *   CALENDAR_ID
 *   SPREADSHEET_ID
 *   SHEET_NAME
 *   ROSTER_COLUMN
 *   ROSTER_START_ROW
 *   DATE_HEADER_ROW
 *   PRESENT_PASS_HOUR
 *   PRESENT_PASS_MINUTE
 *   FINAL_PASS_HOUR
 *   FINAL_PASS_MINUTE
 *   TIMEZONE
 *   STATUS_PRESENT
 *   STATUS_LATE
 *   STATUS_ABSENT
 *   STATUS_EXCUSED
 *   GEMINI_MODEL
 *   GEMINI_API_KEY
 *
 * Gemini is used only for participant-to-roster name matching.
 * Apps Script determines attendance status using the configured
 * status values stored in Script Properties.
 */

const GEMINI_API_URL =
  'https://generativelanguage.googleapis.com/v1beta/models/';

const MEET_PAGE_SIZE = 250;


/**
 * Production entry points.
 */
function runPresentPass() {
  runAttendancePass_('PRESENT_PASS');
}

function runFinalPass() {
  runAttendancePass_('FINAL_PASS');
}


/**
 * Creates the weekday production triggers.
 * Safe to run repeatedly because existing attendance triggers
 * are deleted first.
 */
function setupProduction() {
  const config = getConfig_(true);

  deleteAttendanceTriggers_();

  const weekdays = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY
  ];

  weekdays.forEach(day => {
    ScriptApp.newTrigger('runPresentPass')
      .timeBased()
      .onWeekDay(day)
      .atHour(config.presentPassHour)
      .nearMinute(config.presentPassMinute)
      .inTimezone(config.timezone)
      .create();

    ScriptApp.newTrigger('runFinalPass')
      .timeBased()
      .onWeekDay(day)
      .atHour(config.finalPassHour)
      .nearMinute(config.finalPassMinute)
      .inTimezone(config.timezone)
      .create();
  });

  Logger.log(
    'Production attendance triggers installed successfully.'
  );
}

function deleteAttendanceTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();

    if (
      handler === 'runPresentPass' ||
      handler === 'runFinalPass'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('Existing attendance triggers removed.');
}


/**
 * Core attendance workflow.
 */
function runAttendancePass_(passType) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error(
      'Could not obtain the script lock. Another attendance execution may already be running.'
    );
  }

  try {
    const config = getConfig_(true);
    const now = new Date();

    Logger.log(
      'Starting attendance pass: ' +
        passType +
        ' at ' +
        now.toISOString()
    );

    const event = getTodaysRecurringEvent_(now, config);
    const meetingCode = extractMeetingCode_(event);

    const conferenceRecord =
      getTodaysConferenceRecord_(meetingCode, now);

    const participantNames =
      passType === 'PRESENT_PASS'
        ? getActiveParticipantNames_(conferenceRecord.name)
        : getAllParticipantNames_(conferenceRecord.name);

    Logger.log(
      'Participant records found: ' +
        participantNames.length
    );

    if (
      passType === 'PRESENT_PASS' &&
      participantNames.length === 0
    ) {
      Logger.log(
        'No active participants found. Spreadsheet will remain unchanged.'
      );
      return;
    }

    const sheet = getAttendanceSheet_(config);
    const roster = getRoster_(sheet, config);

    if (roster.length === 0) {
      throw new Error(
        'No roster names were found in the configured roster range.'
      );
    }

    const dateColumn = findDateColumn_(
      sheet,
      now,
      config
    );

    if (dateColumn === -1) {
      throw new Error(
        'No existing attendance column was found for ' +
          formatAttendanceDate_(now, config) +
          '. Spreadsheet was not modified.'
      );
    }

    const matchedRosterNames =
      matchParticipantsWithGemini_(
        participantNames,
        roster,
        config
      );

    Logger.log(
      'Gemini matched roster records: ' +
        matchedRosterNames.size
    );

    if (passType === 'PRESENT_PASS') {
      applyPresentPass_(
        sheet,
        roster,
        dateColumn,
        matchedRosterNames,
        config
      );
    } else {
      applyFinalPass_(
        sheet,
        roster,
        dateColumn,
        matchedRosterNames,
        config
      );
    }

    Logger.log(
      'Attendance pass completed successfully.'
    );
  } catch (error) {
    Logger.log(
      'ATTENDANCE ERROR: ' +
        error.message
    );

    throw error;
  } finally {
    lock.releaseLock();
  }
}


/**
 * PRESENT pass:
 * - Matched active participants -> Present
 * - Existing Present stays Present
 * - Existing Excused stays Excused
 * - Everything else is left unchanged
 */
function applyPresentPass_(
  sheet,
  roster,
  dateColumn,
  matchedRosterNames,
  config
) {
  const range = sheet.getRange(
    config.rosterStartRow,
    dateColumn,
    roster.length,
    1
  );

  const values = range.getValues();

  const updatedValues = roster.map(
    (rosterName, index) => {
      const currentStatus = String(
        values[index][0] || ''
      ).trim();

      if (currentStatus === config.statusExcused) {
        return [config.statusExcused];
      }

      if (currentStatus === config.statusPresent) {
        return [config.statusPresent];
      }

      if (
        matchedRosterNames.has(
          normalizeName_(rosterName)
        )
      ) {
        return [config.statusPresent];
      }

      return [currentStatus];
    }
  );

  range.setValues(updatedValues);

  Logger.log('PRESENT pass applied.');
}


/**
 * FINAL pass:
 * - Existing Present stays Present
 * - Existing Excused stays Excused
 * - Attended but not Present -> Late
 * - Everyone else -> Unexcused
 */
function applyFinalPass_(
  sheet,
  roster,
  dateColumn,
  matchedRosterNames,
  config
) {
  const range = sheet.getRange(
    config.rosterStartRow,
    dateColumn,
    roster.length,
    1
  );

  const values = range.getValues();

  const updatedValues = roster.map(
    (rosterName, index) => {
      const currentStatus = String(
        values[index][0] || ''
      ).trim();

      if (currentStatus === config.statusExcused) {
        return [config.statusExcused];
      }

      if (currentStatus === config.statusPresent) {
        return [config.statusPresent];
      }

      if (
        matchedRosterNames.has(
          normalizeName_(rosterName)
        )
      ) {
        return [config.statusLate];
      }

      return [config.statusAbsent];
    }
  );

  range.setValues(updatedValues);

  Logger.log('FINAL pass applied.');
}


/**
 * Calendar.
 */
function getTodaysRecurringEvent_(now, config) {
  const dayStart = new Date(now);
  dayStart.setHours(0, 0, 0, 0);

  const dayEnd = new Date(now);
  dayEnd.setHours(23, 59, 59, 999);

  const response = Calendar.Events.instances(
    config.calendarId,
    config.recurringEventId,
    {
      timeMin: dayStart.toISOString(),
      timeMax: dayEnd.toISOString(),
      showDeleted: false,
      maxResults: 10
    }
  );

  const items = response.items || [];

  if (items.length === 0) {
    throw new Error(
      'No occurrence of the configured recurring Calendar event was found for today.'
    );
  }

  const todayKey = Utilities.formatDate(
    now,
    config.timezone,
    'yyyy-MM-dd'
  );

  const matchingEvent = items.find(event => {
    const start = getEventStartDate_(event);

    if (!start) {
      return false;
    }

    return (
      Utilities.formatDate(
        start,
        config.timezone,
        'yyyy-MM-dd'
      ) === todayKey
    );
  });

  if (!matchingEvent) {
    throw new Error(
      'The recurring Calendar event was found, but it has no occurrence for today.'
    );
  }

  return matchingEvent;
}

function getEventStartDate_(event) {
  if (event.start && event.start.dateTime) {
    return new Date(event.start.dateTime);
  }

  if (event.start && event.start.date) {
    return new Date(event.start.date + 'T00:00:00');
  }

  return null;
}

function extractMeetingCode_(event) {
  if (
    event.conferenceData &&
    event.conferenceData.entryPoints
  ) {
    const videoEntry =
      event.conferenceData.entryPoints.find(
        entryPoint =>
          entryPoint.entryPointType === 'video'
      );

    if (videoEntry) {
      if (videoEntry.meetingCode) {
        return videoEntry.meetingCode;
      }

      if (videoEntry.uri) {
        const match = videoEntry.uri.match(
          /meet\.google\.com\/([a-z0-9-]+)/i
        );

        if (match) {
          return match[1];
        }
      }
    }
  }

  if (event.hangoutLink) {
    const match = event.hangoutLink.match(
      /meet\.google\.com\/([a-z0-9-]+)/i
    );

    if (match) {
      return match[1];
    }
  }

  throw new Error(
    'No Google Meet code was found in the Calendar event.'
  );
}


/**
 * Google Meet API.
 */
function getTodaysConferenceRecord_(
  meetingCode,
  now
) {
  const dayStart = new Date(now);
  dayStart.setHours(0, 0, 0, 0);

  const dayEnd = new Date(now);
  dayEnd.setHours(23, 59, 59, 999);

  const filter =
    'space.meeting_code = "' +
    meetingCode +
    '"' +
    ' AND start_time >= "' +
    dayStart.toISOString() +
    '"' +
    ' AND start_time <= "' +
    dayEnd.toISOString() +
    '"';

  const response = meetApiRequest_(
    '/conferenceRecords',
    {
      pageSize: 100,
      filter: filter
    }
  );

  const records =
    response.conferenceRecords || [];

  if (records.length === 0) {
    throw new Error(
      'No Meet conference record was found for today.'
    );
  }

  const eligible = records.filter(record => {
    if (!record.startTime) {
      return true;
    }

    return new Date(record.startTime) <= now;
  });

  if (eligible.length === 0) {
    return records[0];
  }

  eligible.sort(
    (a, b) =>
      new Date(b.startTime || 0) -
      new Date(a.startTime || 0)
  );

  return eligible[0];
}

function getActiveParticipantNames_(
  conferenceRecordName
) {
  return listParticipantNames_(
    conferenceRecordName,
    'latest_end_time IS NULL'
  );
}

function getAllParticipantNames_(
  conferenceRecordName
) {
  return listParticipantNames_(
    conferenceRecordName,
    null
  );
}

function listParticipantNames_(
  conferenceRecordName,
  filter
) {
  let pageToken = null;
  const names = [];

  do {
    const queryParams = {
      pageSize: MEET_PAGE_SIZE
    };

    if (filter) {
      queryParams.filter = filter;
    }

    if (pageToken) {
      queryParams.pageToken = pageToken;
    }

    const response = meetApiRequest_(
      '/' +
        conferenceRecordName +
        '/participants',
      queryParams
    );

    const participants =
      response.participants || [];

    participants.forEach(participant => {
      const displayName =
        getParticipantDisplayName_(
          participant
        );

      if (displayName) {
        names.push(displayName);
      }
    });

    pageToken =
      response.nextPageToken || null;
  } while (pageToken);

  return [...new Set(names)];
}

function getParticipantDisplayName_(
  participant
) {
  if (
    participant.signedinUser &&
    participant.signedinUser.displayName
  ) {
    return participant.signedinUser
      .displayName
      .trim();
  }

  if (
    participant.anonymousUser &&
    participant.anonymousUser.displayName
  ) {
    return participant.anonymousUser
      .displayName
      .trim();
  }

  if (
    participant.phoneUser &&
    participant.phoneUser.displayName
  ) {
    return participant.phoneUser
      .displayName
      .trim();
  }

  return null;
}

function meetApiRequest_(
  path,
  queryParams
) {
  const accessToken =
    ScriptApp.getOAuthToken();

  let url =
    'https://meet.googleapis.com/v2' +
    path;

  const params = [];

  Object.keys(queryParams || {}).forEach(key => {
    const value = queryParams[key];

    if (
      value !== null &&
      value !== undefined
    ) {
      params.push(
        encodeURIComponent(key) +
          '=' +
          encodeURIComponent(value)
      );
    }
  });

  if (params.length > 0) {
    url += '?' + params.join('&');
  }

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'get',
        headers: {
          Authorization:
            'Bearer ' +
            accessToken
        },
        muteHttpExceptions: true
      }
    );

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'Google Meet API error ' +
        status +
        ': ' +
        body
    );
  }

  return JSON.parse(body);
}


/**
 * Gemini.
 */
function matchParticipantsWithGemini_(
  participantNames,
  roster,
  config
) {
  if (participantNames.length === 0) {
    return new Set();
  }

  const prompt =
    buildGeminiPrompt_(
      participantNames,
      roster
    );

  const url =
    GEMINI_API_URL +
    config.geminiModel +
    ':generateContent';

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0,
      responseMimeType: 'application/json'
    }
  };

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method: 'post',
        contentType:
          'application/json',
        headers: {
          'x-goog-api-key':
            config.geminiApiKey
        },
        payload:
          JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (status < 200 || status >= 300) {
    if (status === 429) {
      throw new Error(
        'Gemini API quota/rate limit reached (HTTP 429). ' +
          'No attendance decision was made.'
      );
    }

    throw new Error(
      'Gemini API error ' +
        status +
        ': ' +
        body
    );
  }

  const json = JSON.parse(body);

  const responseText =
    json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;

  if (!responseText) {
    throw new Error(
      'Gemini returned no usable response.'
    );
  }

  const result =
    parseGeminiJson_(
      responseText
    );

  if (
    !result ||
    !Array.isArray(result.matches)
  ) {
    throw new Error(
      'Gemini response does not contain a valid matches array.'
    );
  }

  const matched = new Set();

  result.matches.forEach(match => {
    if (
      !match ||
      !Number.isInteger(
        match.rosterIndex
      )
    ) {
      return;
    }

    const index =
      match.rosterIndex;

    if (
      index < 0 ||
      index >= roster.length
    ) {
      return;
    }

    const rosterName =
      roster[index];

    if (
      match.rosterName &&
      normalizeName_(
        match.rosterName
      ) ===
        normalizeName_(
          rosterName
        )
    ) {
      matched.add(
        normalizeName_(
          rosterName
        )
      );
    }
  });

  return matched;
}

function buildGeminiPrompt_(
  participantNames,
  roster
) {
  return `
You are a deterministic Google Meet attendance name-matching system.

Your ONLY task is to match Google Meet participant display names to people in the official roster.

The participant display names may differ from the official roster because of:
- capitalization differences
- punctuation differences
- spacing differences
- missing middle names
- missing middle initials
- abbreviated names
- different name order
- minor spelling differences
- common display-name variations

Rules:
1. Match only to a person that actually exists in the official roster.
2. Never invent a person.
3. Never create a new roster name.
4. Never match one participant to multiple roster entries.
5. If a participant cannot be reliably matched, do not include a match.
6. Use contextual name reasoning where appropriate.
7. When returning rosterName, copy the roster name exactly.
8. rosterIndex is zero-based.
9. Do not decide attendance status. Apps Script applies the configured attendance statuses.
10. Return valid JSON only.

Required output format:
{
  "matches": [
    {
      "participantName": "exact Google Meet participant name",
      "rosterIndex": 0,
      "rosterName": "exact official roster name"
    }
  ]
}

GOOGLE MEET PARTICIPANTS:
${JSON.stringify(participantNames, null, 2)}

OFFICIAL ROSTER:
${JSON.stringify(
  roster.map((name, index) => ({
    index: index,
    name: name
  })),
  null,
  2
)}
`.trim();
}

function parseGeminiJson_(text) {
  let cleaned = text.trim();

  if (cleaned.startsWith('```')) {
    cleaned = cleaned
      .replace(/^```(?:json)?/i, '')
      .replace(/```$/, '')
      .trim();
  }

  return JSON.parse(cleaned);
}


/**
 * Spreadsheet.
 */
function getAttendanceSheet_(config) {
  const spreadsheet =
    SpreadsheetApp.openById(
      config.spreadsheetId
    );

  const sheet =
    spreadsheet.getSheetByName(
      config.sheetName
    );

  if (!sheet) {
    throw new Error(
      'Sheet "' +
        config.sheetName +
        '" does not exist.'
    );
  }

  return sheet;
}

function getRoster_(
  sheet,
  config
) {
  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <
    config.rosterStartRow
  ) {
    return [];
  }

  const numberOfRows =
    lastRow -
    config.rosterStartRow +
    1;

  return sheet
    .getRange(
      config.rosterStartRow,
      config.rosterColumn,
      numberOfRows,
      1
    )
    .getValues()
    .map(row =>
      String(
        row[0] || ''
      ).trim()
    )
    .filter(Boolean);
}

function findDateColumn_(
  sheet,
  date,
  config
) {
  const lastColumn =
    sheet.getLastColumn();

  if (lastColumn < 1) {
    return -1;
  }

  const headers =
    sheet
      .getRange(
        config.dateHeaderRow,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];

  const targetDate =
    formatAttendanceDate_(
      date,
      config
    );

  for (
    let index = 0;
    index < headers.length;
    index++
  ) {
    const header =
      String(
        headers[index] || ''
      )
        .trim()
        .replace(/\s+/g, ' ')
        .toUpperCase();

    if (header === targetDate) {
      return index + 1;
    }
  }

  return -1;
}

function formatAttendanceDate_(
  date,
  config
) {
  return Utilities.formatDate(
    date,
    config.timezone,
    'MMM d'
  ).toUpperCase();
}


/**
 * Configuration.
 */
function getConfig_(
  requireGeminiKey
) {
  const properties =
    PropertiesService.getScriptProperties();

  const getRequired =
    key => {
      const value =
        properties.getProperty(key);

      if (
        value === null ||
        value.trim() === ''
      ) {
        throw new Error(
          'Missing Script Property: ' +
            key
        );
      }

      return value.trim();
    };

  const parsePositiveInteger =
    key => {
      const value =
        Number(
          getRequired(key)
        );

      if (
        !Number.isInteger(value) ||
        value < 1
      ) {
        throw new Error(
          'Script Property ' +
            key +
            ' must be a positive integer.'
        );
      }

      return value;
    };

  const parseHour =
    key => {
      const value =
        Number(
          getRequired(key)
        );

      if (
        !Number.isInteger(value) ||
        value < 0 ||
        value > 23
      ) {
        throw new Error(
          'Script Property ' +
            key +
            ' must be an hour from 0 to 23.'
        );
      }

      return value;
    };

  const parseMinute =
    key => {
      const value =
        Number(
          getRequired(key)
        );

      if (
        !Number.isInteger(value) ||
        value < 0 ||
        value > 59
      ) {
        throw new Error(
          'Script Property ' +
            key +
            ' must be a minute from 0 to 59.'
        );
      }

      return value;
    };

  const config = {
    recurringEventId:
      getRequired('RECURRING_EVENT_ID'),

    calendarId:
      getRequired('CALENDAR_ID'),

    spreadsheetId:
      getRequired('SPREADSHEET_ID'),

    sheetName:
      getRequired('SHEET_NAME'),

    rosterColumn:
      parsePositiveInteger(
        'ROSTER_COLUMN'
      ),

    rosterStartRow:
      parsePositiveInteger(
        'ROSTER_START_ROW'
      ),

    dateHeaderRow:
      parsePositiveInteger(
        'DATE_HEADER_ROW'
      ),

    presentPassHour:
      parseHour(
        'PRESENT_PASS_HOUR'
      ),

    presentPassMinute:
      parseMinute(
        'PRESENT_PASS_MINUTE'
      ),

    finalPassHour:
      parseHour(
        'FINAL_PASS_HOUR'
      ),

    finalPassMinute:
      parseMinute(
        'FINAL_PASS_MINUTE'
      ),

    statusPresent:
      getRequired('STATUS_PRESENT'),

    statusLate:
      getRequired('STATUS_LATE'),

    statusAbsent:
      getRequired('STATUS_ABSENT'),

    statusExcused:
      getRequired('STATUS_EXCUSED'),

    timezone:
      getRequired('TIMEZONE'),

    geminiModel:
      getRequired('GEMINI_MODEL'),

    geminiApiKey:
      requireGeminiKey
        ? getRequired('GEMINI_API_KEY')
        : null
  };

  return config;
}


/**
 * Utility.
 */
function normalizeName_(name) {
  return String(name || '')
    .normalize('NFKC')
    .trim()
    .toLowerCase()
    .replace(/[.,'’`"-]/g, '')
    .replace(/\s+/g, ' ');
}

function columnToLetter_(column) {
  let temp = column;
  let letter = '';

  while (temp > 0) {
    const remainder =
      (temp - 1) % 26;

    letter =
      String.fromCharCode(
        65 + remainder
      ) + letter;

    temp =
      Math.floor(
        (temp - 1) / 26
      );
  }

  return letter;
}


/**
 * Manual tests.
 */
function listTodaysCalendarEvents() {
  const config =
    getConfig_(false);

  const now = new Date();

  const start =
    new Date(now);
  start.setHours(0, 0, 0, 0);

  const end =
    new Date(now);
  end.setHours(23, 59, 59, 999);

  const response =
    Calendar.Events.list(
      config.calendarId,
      {
        timeMin:
          start.toISOString(),
        timeMax:
          end.toISOString(),
        singleEvents:
          true,
        showDeleted:
          false,
        maxResults:
          100
      }
    );

  const events =
    response.items || [];

  events.forEach(event => {
    Logger.log(
      '\n' +
      'SUMMARY: ' +
        event.summary +
      '\n' +
      'ID: ' +
        event.id +
      '\n' +
      'START: ' +
        JSON.stringify(
          event.start
        ) +
      '\n' +
      'RECURRING EVENT ID: ' +
        (
          event.recurringEventId ||
          'NONE'
        ) +
      '\n' +
      'MEET LINK: ' +
        (
          event.hangoutLink ||
          'NONE'
        )
    );
  });

  Logger.log(
    'Total events found: ' +
      events.length
  );
}

function testCalendarAndMeetLookup() {
  const config =
    getConfig_(false);

  const now = new Date();

  const event =
    getTodaysRecurringEvent_(
      now,
      config
    );

  Logger.log(
    'Summary: ' +
      event.summary
  );

  Logger.log(
    'Event ID: ' +
      event.id
  );

  Logger.log(
    'Recurring Event ID: ' +
      (
        event.recurringEventId ||
        'NONE'
      )
  );

  Logger.log(
    'Meet code: ' +
      extractMeetingCode_(event)
  );
}

function testGetCurrentParticipants() {
  const config =
    getConfig_(false);

  const now = new Date();

  const event =
    getTodaysRecurringEvent_(
      now,
      config
    );

  const meetingCode =
    extractMeetingCode_(event);

  const record =
    getTodaysConferenceRecord_(
      meetingCode,
      now
    );

  const participants =
    getActiveParticipantNames_(
      record.name
    );

  Logger.log(
    'Active participants:'
  );

  participants.forEach(name =>
    Logger.log('- ' + name)
  );

  Logger.log(
    'Count: ' +
      participants.length
  );
}

function testGetAllParticipants() {
  const config =
    getConfig_(false);

  const now = new Date();

  const event =
    getTodaysRecurringEvent_(
      now,
      config
    );

  const meetingCode =
    extractMeetingCode_(event);

  const record =
    getTodaysConferenceRecord_(
      meetingCode,
      now
    );

  const participants =
    getAllParticipantNames_(
      record.name
    );

  Logger.log(
    'All participants:'
  );

  participants.forEach(name =>
    Logger.log('- ' + name)
  );

  Logger.log(
    'Count: ' +
      participants.length
  );
}

function testGeminiMatching() {
  const config =
    getConfig_(true);

  const now = new Date();

  const sheet =
    getAttendanceSheet_(config);

  const roster =
    getRoster_(
      sheet,
      config
    );

  const event =
    getTodaysRecurringEvent_(
      now,
      config
    );

  const meetingCode =
    extractMeetingCode_(event);

  const record =
    getTodaysConferenceRecord_(
      meetingCode,
      now
    );

  const participants =
    getAllParticipantNames_(
      record.name
    );

  const matched =
    matchParticipantsWithGemini_(
      participants,
      roster,
      config
    );

  Logger.log(
    'Matched roster names:'
  );

  matched.forEach(name =>
    Logger.log('- ' + name)
  );

  Logger.log(
    'Matched count: ' +
      matched.size
  );
}

function testPresentPassNow() {
  runPresentPass();
}

function testFinalPassNow() {
  runFinalPass();
}
