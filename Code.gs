const CONFIG = {
  // ----------------------------
  // Calendar / Meeting
  // ----------------------------

  CALENDAR_ID: 'YOUR_CALENDAR_ID',

  // ID of the recurring Calendar event.
  RECURRING_EVENT_ID: 'YOUR_RECURRING_EVENT_ID',


  // ----------------------------
  // Spreadsheet
  // ----------------------------

  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID',

  SHEET_NAME: 'YOUR_SHEET_NAME',

  // 1-based column number containing roster names.
  // A = 1, B = 2, C = 3, etc.
  ROSTER_COLUMN: YOUR_ROSTER_COLUMN,

  // 1-based row containing attendance dates.
  DATE_HEADER_ROW: YOUR_DATE_HEADER_ROW,

  // First row containing roster names.
  ROSTER_START_ROW: YOUR_ROSTER_START_ROW,


  // ================================
  // ATTENDANCE PASSES
  // ================================

  // PRESENT PASS
  PRESENT_PASS_HOUR: YOUR_PRESENT_PASS_HOUR,
  PRESENT_PASS_MINUTE: YOUR_PRESENT_PASS_MINUTE,

  // FINAL PASS
  FINAL_PASS_HOUR: YOUR_FINAL_PASS_HOUR,
  FINAL_PASS_MINUTE: YOUR_FINAL_PASS_MINUTE,


  // ----------------------------
  // Attendance statuses
  // ----------------------------

  STATUS_PRESENT: 'Present',
  STATUS_LATE: 'Late',
  STATUS_ABSENT: 'Unexcused',
  STATUS_EXCUSED: 'Excused',


  // ----------------------------
  // Gemini
  // ----------------------------

  GEMINI_MODEL: 'YOUR_GEMINI_MODEL',

  GEMINI_API_URL:
    'https://generativelanguage.googleapis.com/v1beta/models/',


  // ================================
  // GENERAL
  // ================================

  TIMEZONE: 'YOUR_TIMEZONE',
  MEET_PAGE_SIZE: 250
};


/**
 * ============================================================
 * PRODUCTION ENTRY POINTS
 * ============================================================
 */

/**
 * Trigger entry point for the PRESENT pass.
 */
function runPresentPass() {
  runAttendancePass_('PRESENT_PASS');
}


/**
 * Trigger entry point for the FINAL pass.
 */
function runFinalPass() {
  runAttendancePass_('FINAL_PASS');
}


/**
 * ============================================================
 * PRODUCTION SETUP
 * ============================================================
 */

/**
 * Run this ONCE when deploying the automation.
 *
 * It:
 *   1. Validates configuration.
 *   2. Removes previous attendance triggers.
 *   3. Creates the 10 weekday triggers.
 *
 * Running this again is safe because existing attendance
 * triggers are removed first.
 */
function setupProduction() {
  validateConfiguration_();

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
      .atHour(CONFIG.PRESENT_PASS_HOUR)
      .nearMinute(CONFIG.PRESENT_PASS_MINUTE)
      .inTimezone(CONFIG.TIMEZONE)
      .create();

    ScriptApp.newTrigger('runFinalPass')
      .timeBased()
      .onWeekDay(day)
      .atHour(CONFIG.FINAL_PASS_HOUR)
      .nearMinute(CONFIG.FINAL_PASS_MINUTE)
      .inTimezone(CONFIG.TIMEZONE)
      .create();
  });

  Logger.log(
    'Production attendance triggers installed successfully.'
  );
}


/**
 * Removes only triggers belonging to this automation.
 */
function deleteAttendanceTriggers_() {
  const triggers =
    ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    const handler =
      trigger.getHandlerFunction();

    if (
      handler === 'runPresentPass' ||
      handler === 'runFinalPass'
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log(
    'Existing attendance triggers removed.'
  );
}


/**
 * ============================================================
 * CORE ATTENDANCE WORKFLOW
 * ============================================================
 */

function runAttendancePass_(passType) {
  const lock =
    LockService.getScriptLock();

  /**
   * Prevent simultaneous duplicate execution.
   */
  if (!lock.tryLock(30000)) {
    throw new Error(
      'Could not obtain the script lock. Another attendance execution may already be running.'
    );
  }

  try {
    validateConfiguration_();

    const now = new Date();

    Logger.log(
      'Starting attendance pass: ' +
        passType
    );

    Logger.log(
      'Execution time: ' +
        now.toISOString()
    );


    // ------------------------------------
    // 1. Find today's Calendar occurrence
    // ------------------------------------

    const event =
      getTodaysRecurringEvent_(now);

    Logger.log(
      'Calendar event found: ' +
        event.summary
    );


    // ------------------------------------
    // 2. Get Meet code
    // ------------------------------------

    const meetingCode =
      extractMeetingCode_(event);

    Logger.log(
      'Meet code resolved successfully.'
    );


    // ------------------------------------
    // 3. Get today's Meet conference
    // ------------------------------------

    const conferenceRecord =
      getTodaysConferenceRecord_(
        meetingCode,
        now
      );

    Logger.log(
      'Meet conference record found.'
    );


    // ------------------------------------
    // 4. Get participant names
    // ------------------------------------

    let participantNames;

    if (passType === 'PRESENT_PASS') {
      participantNames =
        getActiveParticipantNames_(
          conferenceRecord.name
        );
    } else {
      participantNames =
        getAllParticipantNames_(
          conferenceRecord.name
        );
    }

    Logger.log(
      'Participant records found: ' +
        participantNames.length
    );


    // ------------------------------------
    // 5. 1:30 "nobody present" rule
    // ------------------------------------

    if (
      passType === 'PRESENT_PASS' &&
      participantNames.length === 0
    ) {
      Logger.log(
        'No active participants found. ' +
          'Spreadsheet will remain unchanged.'
      );

      return;
    }


    // ------------------------------------
    // 6. Get roster
    // ------------------------------------

    const sheet =
      getAttendanceSheet_();

    const roster =
      getRoster_(sheet);

    if (roster.length === 0) {
      throw new Error(
        'The roster in column A is empty.'
      );
    }

    Logger.log(
      'Roster size: ' +
        roster.length
    );


    // ------------------------------------
    // 7. Find today's date column
    // ------------------------------------

    const dateColumn =
      findDateColumn_(
        sheet,
        now
      );

    if (dateColumn === -1) {
      throw new Error(
        'No existing attendance column was found for ' +
          formatAttendanceDate_(now) +
          '. Spreadsheet was not modified.'
      );
    }

    Logger.log(
      'Attendance column: ' +
        columnToLetter_(dateColumn)
    );


    // ------------------------------------
    // 8. Gemini name matching
    // ------------------------------------

    const matchedRosterNames =
      matchParticipantsWithGemini_(
        participantNames,
        roster
      );

    Logger.log(
      'Gemini matched roster records: ' +
        matchedRosterNames.size
    );


    // ------------------------------------
    // 9. Deterministic attendance logic
    // ------------------------------------

    if (passType === 'PRESENT_PASS') {
      applyPresentPass_(
        sheet,
        roster,
        dateColumn,
        matchedRosterNames
      );
    } else {
      applyFinalPass_(
        sheet,
        roster,
        dateColumn,
        matchedRosterNames
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
 * ============================================================
 * PRESENT PASS
 * ============================================================
 */

function applyPresentPass_(
  sheet,
  roster,
  dateColumn,
  matchedRosterNames
) {
  const range = sheet.getRange(
    CONFIG.ROSTER_START_ROW,
    dateColumn,
    roster.length,
    1
  );

  const values =
    range.getValues();

  const updatedValues =
    roster.map((rosterName, index) => {

      const currentStatus =
        String(
          values[index][0] || ''
        ).trim();


      // Never overwrite EXCUSED.
      if (
        currentStatus ===
        CONFIG.STATUS_EXCUSED
      ) {
        return [
          CONFIG.STATUS_EXCUSED
        ];
      }


      // Existing PRESENT remains PRESENT.
      if (
        currentStatus ===
        CONFIG.STATUS_PRESENT
      ) {
        return [
          CONFIG.STATUS_PRESENT
        ];
      }


      // Gemini identified this roster member
      // as currently present.
      if (
        matchedRosterNames.has(
          normalizeName_(
            rosterName
          )
        )
      ) {
        return [
          CONFIG.STATUS_PRESENT
        ];
      }


      // Leave everything else untouched
      // during the 1:30 pass.
      return [
        currentStatus
      ];
    });

  range.setValues(
    updatedValues
  );

  Logger.log(
    'PRESENT pass applied.'
  );
}


/**
 * ============================================================
 * FINAL PASS
 * ============================================================
 */

function applyFinalPass_(
  sheet,
  roster,
  dateColumn,
  matchedRosterNames
) {
  const range = sheet.getRange(
    CONFIG.ROSTER_START_ROW,
    dateColumn,
    roster.length,
    1
  );

  const values =
    range.getValues();

  const updatedValues =
    roster.map((rosterName, index) => {

      const currentStatus =
        String(
          values[index][0] || ''
        ).trim();


      // EXCUSED is manually managed.
      // Never overwrite it.
      if (
        currentStatus ===
        CONFIG.STATUS_EXCUSED
      ) {
        return [
          CONFIG.STATUS_EXCUSED
        ];
      }


      // Existing PRESENT remains PRESENT.
      if (
        currentStatus ===
        CONFIG.STATUS_PRESENT
      ) {
        return [
          CONFIG.STATUS_PRESENT
        ];
      }


      // Attended, but was not present at
      // the 1:30 pass.
      if (
        matchedRosterNames.has(
          normalizeName_(
            rosterName
          )
        )
      ) {
        return [
          CONFIG.STATUS_LATE
        ];
      }


      // No attendance record.
      return [
        CONFIG.STATUS_ABSENT
      ];
    });

  range.setValues(
    updatedValues
  );

  Logger.log(
    'FINAL pass applied.'
  );
}


/**
 * ============================================================
 * CALENDAR
 * ============================================================
 */

function getTodaysRecurringEvent_(now) {
  const dayStart =
    new Date(now);

  dayStart.setHours(
    0,
    0,
    0,
    0
  );


  const dayEnd =
    new Date(now);

  dayEnd.setHours(
    23,
    59,
    59,
    999
  );


  const response =
    Calendar.Events.instances(
      CONFIG.CALENDAR_ID,
      CONFIG.RECURRING_EVENT_ID,
      {
        timeMin:
          dayStart.toISOString(),

        timeMax:
          dayEnd.toISOString(),

        showDeleted:
          false,

        maxResults:
          10
      }
    );


  const items =
    response.items || [];


  if (items.length === 0) {
    throw new Error(
      'No occurrence of the configured recurring Calendar event was found for today.'
    );
  }


  const todayKey =
    Utilities.formatDate(
      now,
      CONFIG.TIMEZONE,
      'yyyy-MM-dd'
    );


  const matchingEvent =
    items.find(event => {

      const start =
        getEventStartDate_(
          event
        );

      if (!start) {
        return false;
      }

      return (
        Utilities.formatDate(
          start,
          CONFIG.TIMEZONE,
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
  if (
    event.start &&
    event.start.dateTime
  ) {
    return new Date(
      event.start.dateTime
    );
  }


  if (
    event.start &&
    event.start.date
  ) {
    return new Date(
      event.start.date +
        'T00:00:00'
    );
  }


  return null;
}


/**
 * Get the Meet code from the Calendar event.
 */
function extractMeetingCode_(event) {
  if (
    event.conferenceData &&
    event.conferenceData.entryPoints
  ) {
    const videoEntry =
      event.conferenceData.entryPoints.find(
        entryPoint =>
          entryPoint.entryPointType ===
          'video'
      );


    if (videoEntry) {

      if (
        videoEntry.meetingCode
      ) {
        return videoEntry.meetingCode;
      }


      if (
        videoEntry.uri
      ) {
        const match =
          videoEntry.uri.match(
            /meet\.google\.com\/([a-z0-9-]+)/i
          );

        if (match) {
          return match[1];
        }
      }
    }
  }


  if (event.hangoutLink) {
    const match =
      event.hangoutLink.match(
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
 * ============================================================
 * MEET API
 * ============================================================
 */

function getTodaysConferenceRecord_(
  meetingCode,
  now
) {
  const dayStart =
    new Date(now);

  dayStart.setHours(
    0,
    0,
    0,
    0
  );


  const dayEnd =
    new Date(now);

  dayEnd.setHours(
    23,
    59,
    59,
    999
  );


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


  const response =
    meetApiRequest_(
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


  const eligible =
    records.filter(record => {

      if (!record.startTime) {
        return true;
      }

      return (
        new Date(
          record.startTime
        ) <= now
      );
    });


  if (
    eligible.length === 0
  ) {
    return records[0];
  }


  eligible.sort(
    (a, b) =>
      new Date(
        b.startTime || 0
      ) -
      new Date(
        a.startTime || 0
      )
  );


  return eligible[0];
}


/**
 * Get participants currently active.
 *
 * Google documents:
 *
 * latest_end_time IS NULL
 *
 * as the filter for active participants.
 */
function getActiveParticipantNames_(
  conferenceRecordName
) {
  return listParticipantNames_(
    conferenceRecordName,
    'latest_end_time IS NULL'
  );
}


/**
 * Get every participant who attended.
 */
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
  let pageToken =
    null;

  const names = [];


  do {

    const queryParams = {
      pageSize:
        CONFIG.MEET_PAGE_SIZE
    };


    if (filter) {
      queryParams.filter =
        filter;
    }


    if (pageToken) {
      queryParams.pageToken =
        pageToken;
    }


    const response =
      meetApiRequest_(
        '/' +
          conferenceRecordName +
          '/participants',
        queryParams
      );


    const participants =
      response.participants || [];


    participants.forEach(
      participant => {

        const displayName =
          getParticipantDisplayName_(
            participant
          );


        if (displayName) {
          names.push(
            displayName
          );
        }
      }
    );


    pageToken =
      response.nextPageToken ||
      null;

  } while (pageToken);


  return [
    ...new Set(names)
  ];
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


  Object.keys(
    queryParams || {}
  ).forEach(key => {

    const value =
      queryParams[key];


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


  if (
    params.length > 0
  ) {
    url +=
      '?' +
      params.join('&');
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

        muteHttpExceptions:
          true
      }
    );


  const status =
    response.getResponseCode();

  const body =
    response.getContentText();


  if (
    status < 200 ||
    status >= 300
  ) {
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
 * ============================================================
 * GEMINI
 * ============================================================
 */

function matchParticipantsWithGemini_(
  participantNames,
  roster
) {
  if (
    participantNames.length === 0
  ) {
    return new Set();
  }


  const apiKey =
    PropertiesService
      .getScriptProperties()
      .getProperty(
        'GEMINI_API_KEY'
      );


  if (!apiKey) {
    throw new Error(
      'GEMINI_API_KEY is not configured in Script Properties.'
    );
  }


  const prompt =
    buildGeminiPrompt_(
      participantNames,
      roster
    );


  const url =
    CONFIG.GEMINI_API_URL +
    CONFIG.GEMINI_MODEL +
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

      responseMimeType:
        'application/json'
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
            apiKey
        },

        payload:
          JSON.stringify(payload),

        muteHttpExceptions:
          true
      }
    );


  const status =
    response.getResponseCode();

  const body =
    response.getContentText();


  if (
    status < 200 ||
    status >= 300
  ) {
    throw new Error(
      'Gemini API error ' +
        status +
        ': ' +
        body
    );
  }


  const json =
    JSON.parse(body);


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
    !Array.isArray(
      result.matches
    )
  ) {
    throw new Error(
      'Gemini response does not contain a valid matches array.'
    );
  }


  const matched =
    new Set();


  result.matches.forEach(
    match => {

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


      /**
       * We trust only the roster index.
       *
       * The returned rosterName is checked against
       * the actual roster record before accepting it.
       */
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
    }
  );


  return matched;
}


/**
 * Gemini is deliberately limited to matching.
 *
 * Apps Script remains responsible for attendance status.
 */
function buildGeminiPrompt_(
  participantNames,
  roster
) {
  return `
You are a deterministic Google Meet attendance name-matching system.

Your ONLY task is to match Google Meet participant display names
to people in the official roster.

The participant display names may differ from the official roster
because of:

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
9. Do not decide PRESENT, LATE, ABSENT, or EXCUSED.
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
${JSON.stringify(
  participantNames,
  null,
  2
)}

OFFICIAL ROSTER:
${JSON.stringify(
  roster.map(
    (name, index) => ({
      index: index,
      name: name
    })
  ),
  null,
  2
)}
`.trim();
}


function parseGeminiJson_(
  text
) {
  let cleaned =
    text.trim();


  if (
    cleaned.startsWith('```')
  ) {
    cleaned =
      cleaned
        .replace(
          /^```(?:json)?/i,
          ''
        )
        .replace(
          /```$/,
          ''
        )
        .trim();
  }


  return JSON.parse(
    cleaned
  );
}


/**
 * ============================================================
 * SPREADSHEET
 * ============================================================
 */

function getAttendanceSheet_() {
  const spreadsheet =
    SpreadsheetApp.openById(
      CONFIG.SPREADSHEET_ID
    );


  const sheet =
    spreadsheet.getSheetByName(
      CONFIG.SHEET_NAME
    );


  if (!sheet) {
    throw new Error(
      'Sheet "' +
        CONFIG.SHEET_NAME +
        '" does not exist.'
    );
  }


  return sheet;
}


function getRoster_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.ROSTER_START_ROW) {
    return [];
  }

  const numberOfRows =
    lastRow - CONFIG.ROSTER_START_ROW + 1;

  return sheet
    .getRange(
      CONFIG.ROSTER_START_ROW,
      CONFIG.ROSTER_COLUMN,
      numberOfRows,
      1
    )
    .getValues()
    .map(row => String(row[0] || '').trim())
    .filter(Boolean);
}


/**
 * Find today's existing date column.
 *
 * This function NEVER creates columns.
 */
function findDateColumn_(sheet, date) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn < 1) {
    return -1;
  }

  const headers = sheet
    .getRange(
      CONFIG.DATE_HEADER_ROW,
      1,
      1,
      lastColumn
    )
    .getDisplayValues()[0];

  const targetDate =
    formatAttendanceDate_(date);

  for (let index = 0; index < headers.length; index++) {
    const header = String(headers[index] || '')
      .trim()
      .replace(/\s+/g, ' ')
      .toUpperCase();

    if (header === targetDate) {
      return index + 1;
    }
  }

  return -1;
}


function formatAttendanceDate_(date) {
  return Utilities.formatDate(
    date,
    CONFIG.TIMEZONE,
    'MMM d'
  ).toUpperCase();
}


/**
 * ============================================================
 * NAME NORMALIZATION
 * ============================================================
 */

function normalizeName_(
  name
) {
  return String(
    name || ''
  )
    .normalize('NFKC')
    .trim()
    .toLowerCase()
    .replace(
      /[.,'’`"-]/g,
      ''
    )
    .replace(
      /\s+/g,
      ' '
    );
}


/**
 * ============================================================
 * VALIDATION
 * ============================================================
 */

function validateConfiguration_() {
  if (
    !CONFIG.RECURRING_EVENT_ID ||
    CONFIG.RECURRING_EVENT_ID ===
      'PUT_RECURRING_EVENT_ID_HERE'
  ) {
    throw new Error(
      'RECURRING_EVENT_ID has not been configured.'
    );
  }


  if (
    !CONFIG.SPREADSHEET_ID
  ) {
    throw new Error(
      'SPREADSHEET_ID is missing.'
    );
  }


  if (
    !CONFIG.SHEET_NAME
  ) {
    throw new Error(
      'SHEET_NAME is missing.'
    );
  }


  if (
    !CONFIG.GEMINI_MODEL
  ) {
    throw new Error(
      'GEMINI_MODEL is missing.'
    );
  }
}


/**
 * ============================================================
 * UTILITIES
 * ============================================================
 */

function columnToLetter_(
  column
) {
  let temp =
    column;

  let letter =
    '';

  while (
    temp > 0
  ) {
    const remainder =
      (temp - 1) % 26;

    letter =
      String.fromCharCode(
        65 + remainder
      ) +
      letter;

    temp =
      Math.floor(
        (temp - 1) / 26
      );
  }


  return letter;
}


/**
 * ============================================================
 * MANUAL TEST FUNCTIONS
 * ============================================================
 */

function listTodaysCalendarEvents() {
  const now = new Date();

  const start = new Date(now);
  start.setHours(0, 0, 0, 0);

  const end = new Date(now);
  end.setHours(23, 59, 59, 999);

  const response = Calendar.Events.list(
    CONFIG.CALENDAR_ID,
    {
      timeMin: start.toISOString(),
      timeMax: end.toISOString(),
      singleEvents: true,
      showDeleted: false,
      maxResults: 100
    }
  );

  const events = response.items || [];

  events.forEach(event => {
    Logger.log(
      '\n' +
      'SUMMARY: ' + event.summary + '\n' +
      'ID: ' + event.id + '\n' +
      'START: ' + JSON.stringify(event.start) + '\n' +
      'RECURRING EVENT ID: ' + (event.recurringEventId || 'NONE') + '\n' +
      'MEET LINK: ' + (event.hangoutLink || 'NONE')
    );
  });

  Logger.log('Total events found: ' + events.length);
}

function testCalendarAndMeetLookup() {
  const now =
    new Date();

  const event =
    getTodaysRecurringEvent_(
      now
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
      extractMeetingCode_(
        event
      )
  );
}


function testGetCurrentParticipants() {
  const now =
    new Date();

  const event =
    getTodaysRecurringEvent_(
      now
    );

  const meetingCode =
    extractMeetingCode_(
      event
    );

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

  participants.forEach(
    name =>
      Logger.log(
        '- ' + name
      )
  );

  Logger.log(
    'Count: ' +
      participants.length
  );
}


function testGetAllParticipants() {
  const now =
    new Date();

  const event =
    getTodaysRecurringEvent_(
      now
    );

  const meetingCode =
    extractMeetingCode_(
      event
    );

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

  participants.forEach(
    name =>
      Logger.log(
        '- ' + name
      )
  );

  Logger.log(
    'Count: ' +
      participants.length
  );
}


function testGeminiMatching() {
  const now =
    new Date();

  const sheet =
    getAttendanceSheet_();

  const roster =
    getRoster_(
      sheet
    );

  const event =
    getTodaysRecurringEvent_(
      now
    );

  const meetingCode =
    extractMeetingCode_(
      event
    );

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
      roster
    );

  Logger.log(
    'Matched roster names:'
  );

  matched.forEach(
    name =>
      Logger.log(
        '- ' + name
      )
  );

  Logger.log(
    'Matched count: ' +
      matched.size
  );
}


/**
 * Runs the actual PRESENT logic immediately.
 */
function testPresentPassNow() {
  runPresentPass();
}


/**
 * Runs the actual FINAL logic immediately.
 */
function testFinalPassNow() {
  runFinalPass();
}
