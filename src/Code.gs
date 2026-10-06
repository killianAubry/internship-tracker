/**
 * Internship Application Tracker (Groq edition)
 *
 * DAILY workflow:
 * Gmail candidate filter -> Groq LLM classifier -> Application ID matching ->
 * Applications / Email Updates / Needs Review.
 *
 * Required Script Property:
 *   GROQ_API_KEY = your key from console.groq.com/keys
 *
 * Setup:
 *   1. Put this code into Extensions -> Apps Script.
 *   2. Add GROQ_API_KEY under Project Settings -> Script properties.
 *   3. Run setupTracker() once and authorize access.
 *   4. Run retryFailedLlmRows() once to retry rows previously marked as
 *      API failures (including old "Gemini classification failed" rows).
 *
 * If a run hits the rate limit or time budget, it stops cleanly and the
 * remaining emails are picked up on the next run (safe to run manually).
 */

const CONFIG = {
  SHEET_APPLICATIONS: 'Applications',
  SHEET_UPDATES: 'Email Updates',
  SHEET_REVIEW: 'Needs Review',
  SHEET_PROCESSED: 'Processed', // hidden; every message ID we've classified

  DAILY_TRIGGER_HOUR: 7,

  GMAIL_QUERY:
    'newer_than:365d (' +
      '"thank you for applying" OR ' +
      '"application received" OR ' +
      '"application submitted" OR ' +
      '"we received your application" OR ' +
      '"thank you for your interest" OR ' +
      '"your application" OR ' +
      '"application status" OR ' +
      '"next steps" OR ' +
      'interview OR ' +
      '"phone screen" OR ' +
      '"technical screen" OR ' +
      '"not moving forward" OR ' +
      '"not selected" OR ' +
      '"regret to inform" OR ' +
      'offer' +
    ') ' +
    '-from:(jobalerts-noreply@linkedin.com) ' +
    '-from:(jobs-noreply@linkedin.com)',

  // Smaller excerpt = fewer tokens per call (free tier has tight TPM limits).
  MAX_BODY_CHARS: 3000,

  AUTO_ACCEPT_CONFIDENCE: 0.82,

  // Free-tier models on Groq (check console.groq.com/docs/models):
  //   'openai/gpt-oss-120b' -> best accuracy
  //   'openai/gpt-oss-20b'  -> faster, lighter on token limits
  GROQ_MODEL: 'openai/gpt-oss-120b',

  // Pause between API calls to stay under requests-per-minute limits.
  CALL_SPACING_MS: 2500,

  // Retry attempts on HTTP 429 / 5xx before giving up for this run.
  MAX_RETRIES: 3,

  // Apps Script hard-kills at 6 min; stop cleanly before that.
  MAX_RUNTIME_MS: 270000
};

const FAILURE_PREFIXES = [
  'LLM classification failed:',
  'Gemini classification failed:' // legacy rows
];


/**
 * Run this ONCE.
 */
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheets_(ss);
  ensureHeaders_(ss);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'scanInternshipEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('scanInternshipEmails')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.DAILY_TRIGGER_HOUR)
    .create();

  Logger.log('Tracker configured. One daily scan trigger was created.');
}


/**
 * Main daily scanner.
 */
function scanInternshipEmails() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets_(ss);
  ensureHeaders_(ss);

  const processedMessageIds = getProcessedMessageIds_(ss);
  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 100);

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
        Logger.log('Time budget reached; remaining emails will be handled next run.');
        return;
      }

      const messageId = message.getId();
      if (processedMessageIds.has(messageId)) continue;

      const candidate = buildCandidateEmail_(message, thread);
      if (!candidate) continue;

      let classification;
      try {
        classification = classifyWithGroq_(candidate);
      } catch (err) {
        if (err.rateLimited || err.fatal) {
          // Not the email's fault (rate limit, bad key, missing model):
          // leave it unprocessed and try again next run.
          Logger.log('Stopping this run: ' + err.message);
          return;
        }
        addReview_(ss, candidate, {
          classification: 'UNCERTAIN',
          confidence: 0,
          company: '',
          role: '',
          status: 'Needs Review',
          reason: 'LLM classification failed: ' + err.message
        }, 'Review manually; Groq API call failed.');
        markProcessed_(ss, candidate);
        continue;
      }

      processClassification_(ss, candidate, classification);
      processedMessageIds.add(messageId);

      Utilities.sleep(CONFIG.CALL_SPACING_MS);
    }
  }
}


/**
 * Retry rows that landed in Needs Review because the LLM call failed.
 * Handles both new and legacy ("Gemini ...") failure prefixes.
 */
function retryFailedLlmRows() {
  const startTime = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets_(ss);
  ensureHeaders_(ss);

  const review = ss.getSheetByName(CONFIG.SHEET_REVIEW);
  const values = review.getDataRange().getValues();
  const rowsToDelete = [];

  for (let r = 1; r < values.length; r++) {
    if (Date.now() - startTime > CONFIG.MAX_RUNTIME_MS) {
      Logger.log('Time budget reached; run again to continue.');
      break;
    }

    const reason = String(values[r][11] || ''); // Reason, column L
    if (!FAILURE_PREFIXES.some(p => reason.startsWith(p))) continue;

    const messageId = String(values[r][9] || ''); // Message ID, column J
    if (!messageId) continue;

    try {
      const message = GmailApp.getMessageById(messageId);
      if (!message) continue;

      const email = buildCandidateEmail_(message, message.getThread());
      const classification = classifyWithGroq_(email);
      processClassification_(ss, email, classification);
      rowsToDelete.push(r + 1);
      Utilities.sleep(CONFIG.CALL_SPACING_MS);
    } catch (err) {
      Logger.log('Retry failed for message ' + messageId + ': ' + err.message);
      if (err.rateLimited || err.fatal) break;
    }
  }

  rowsToDelete.sort((a, b) => b - a).forEach(row => review.deleteRow(row));
  Logger.log('Retried ' + rowsToDelete.length + ' previously failed classifications.');
}


function buildCandidateEmail_(message, thread) {
  const subject = message.getSubject() || '';
  const from = message.getFrom() || '';
  const body = (message.getPlainBody() || '').trim();
  const excerpt = body.substring(0, CONFIG.MAX_BODY_CHARS);

  return {
    received: message.getDate(),
    messageId: message.getId(),
    threadId: thread.getId(),
    subject: subject,
    from: from,
    body: excerpt,
    gmailLink: 'https://mail.google.com/mail/u/0/#inbox/' + thread.getId()
  };
}


/**
 * Groq (OpenAI-compatible) chat completion returning strict JSON.
 */
function classifyWithGroq_(email) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY');
  if (!apiKey) {
    const err = new Error('Missing GROQ_API_KEY in Script Properties.');
    err.fatal = true;
    throw err;
  }

  const systemPrompt = `
You classify emails for a student's internship application tracker.

IMPORTANT:
- Do NOT classify ordinary job alerts, newsletters, marketing, recruiting advertisements,
  university career announcements, or unrelated mail as an application.
- An application-related email should refer to a specific job/internship application,
  recruiter interaction, interview process, application status, offer, or rejection.
- Extract the actual hiring company, not the ATS/vendor (for example Workday, Greenhouse,
  Lever, Ashby, SmartRecruiters) when the company is identifiable.
- Extract the actual internship/job role.
- If the email is ambiguous, return UNCERTAIN rather than guessing.
- A later email about the same company and role should be treated as an UPDATE, not a new application.
- Confidence must be a number between 0 and 1.

Respond with ONLY a JSON object matching this schema:
{
  "classification": "APPLICATION_CONFIRMATION | INTERVIEW_REQUEST | RECRUITER_CONTACT | APPLICATION_UPDATE | OFFER | REJECTION | JOB_ALERT | MARKETING | UNRELATED | UNCERTAIN",
  "confidence": 0.0,
  "company": "",
  "role": "",
  "status": "Applied | Application Received | Under Review | Recruiter Contact | Interview | Offer | Rejected | Withdrawn | Needs Review",
  "reason": "",
  "application_id_hint": ""
}
`.trim();

  const userPrompt = `Email:
From: ${email.from}
Subject: ${email.subject}
Date: ${email.received.toISOString()}

Body:
${email.body}`;

  const payload = {
    model: CONFIG.GROQ_MODEL,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0,
    max_completion_tokens: 1500, // includes reasoning tokens
    reasoning_effort: 'low',
    response_format: { type: 'json_object' }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const url = 'https://api.groq.com/openai/v1/chat/completions';
  let lastError = '';

  for (let attempt = 0; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code >= 200 && code < 300) {
      const json = JSON.parse(text);
      const generated = json.choices &&
                        json.choices[0] &&
                        json.choices[0].message &&
                        json.choices[0].message.content;
      if (!generated) {
        throw new Error('Groq returned no usable classification.');
      }
      return JSON.parse(stripMarkdownFence_(generated));
    }

    lastError = 'Groq HTTP ' + code + ': ' + text.substring(0, 300);

    // Retry on rate limit or transient server errors.
    if (code === 429 || code >= 500) {
      const headers = response.getHeaders();
      const retryAfter = Number(headers['retry-after'] || headers['Retry-After'] || 0);
      const waitMs = Math.min(
        Math.max(retryAfter * 1000, 2000 * Math.pow(2, attempt)),
        60000
      );
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(waitMs);
        continue;
      }
      if (code === 429) {
        const err = new Error(lastError);
        err.rateLimited = true;
        throw err;
      }
    }

    // Auth/model problems affect every email, so stop the whole run.
    if (code === 401 || code === 403 || code === 404) {
      const err = new Error(lastError);
      err.fatal = true;
      throw err;
    }

    throw new Error(lastError);
  }

  throw new Error(lastError || 'Groq call failed.');
}


function processClassification_(ss, email, c) {
  const confidence = Number(c.confidence || 0);
  const classification = String(c.classification || 'UNCERTAIN');

  if (['JOB_ALERT', 'MARKETING', 'UNRELATED'].includes(classification)) {
    markProcessed_(ss, email);
    return;
  }

  if (
    classification === 'UNCERTAIN' ||
    confidence < CONFIG.AUTO_ACCEPT_CONFIDENCE ||
    !c.company ||
    !c.role
  ) {
    const existing = findBestApplication_(ss, c.company, c.role, c.application_id_hint);

    addReview_(
      ss,
      email,
      c,
      existing
        ? 'Possible update to ' + existing.applicationId + '; review before applying automatically.'
        : 'Review whether this is a real application and whether it is a new application.'
    );

    markProcessed_(ss, email);
    return;
  }

  const existing = findBestApplication_(ss, c.company, c.role, c.application_id_hint);

  let applicationId;
  if (existing) {
    applicationId = existing.applicationId;
    updateApplication_(ss, existing.row, applicationId, c, email);
  } else {
    applicationId = createApplication_(ss, c, email);
  }

  appendUpdate_(ss, email, c, applicationId);
  markProcessed_(ss, email);
}


function createApplication_(ss, c, email) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_APPLICATIONS);
  const applicationId = nextApplicationId_(sheet);
  const key = makeApplicationKey_(c.company, c.role);

  sheet.appendRow([
    applicationId,
    clean_(c.company),
    clean_(c.role),
    email.received,
    normalizeStatus_(c.status, c.classification),
    c.reason || '',
    email.received,
    email.from,
    email.subject,
    email.gmailLink,
    key,
    '',
    ''
  ]);

  return applicationId;
}


function updateApplication_(ss, row, applicationId, c, email) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_APPLICATIONS);

  if (c.company) sheet.getRange(row, 2).setValue(clean_(c.company));
  if (c.role) sheet.getRange(row, 3).setValue(clean_(c.role));

  sheet.getRange(row, 5).setValue(normalizeStatus_(c.status, c.classification));
  sheet.getRange(row, 6).setValue(c.reason || '');
  sheet.getRange(row, 7).setValue(email.received);
  sheet.getRange(row, 8).setValue(email.from);
  sheet.getRange(row, 9).setValue(email.subject);
  sheet.getRange(row, 10).setValue(email.gmailLink);

  const currentKey = sheet.getRange(row, 11).getValue();
  if (!currentKey) {
    sheet.getRange(row, 11).setValue(makeApplicationKey_(c.company, c.role));
  }
}


/**
 * Priority:
 * 1. Explicit APP-#### hint
 * 2. Normalized company + role
 * 3. Conservative fuzzy role match if the company has exactly one application
 */
function findBestApplication_(ss, company, role, applicationIdHint) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_APPLICATIONS);
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return null;

  const hint = String(applicationIdHint || '').match(/APP-\d{4,}/i);
  if (hint) {
    const wanted = hint[0].toUpperCase();
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][0]).toUpperCase() === wanted) {
        return {
          row: r + 1,
          applicationId: values[r][0],
          company: values[r][1],
          role: values[r][2]
        };
      }
    }
  }

  const companyKey = normalize_(company);
  const roleKey = normalize_(role);

  if (!companyKey || !roleKey) return null;

  for (let r = 1; r < values.length; r++) {
    const existingKey = normalize_(values[r][10]);
    if (existingKey === makeApplicationKey_(company, role)) {
      return {
        row: r + 1,
        applicationId: values[r][0],
        company: values[r][1],
        role: values[r][2]
      };
    }
  }

  const companyMatches = [];
  for (let r = 1; r < values.length; r++) {
    if (normalize_(values[r][1]) === companyKey) {
      companyMatches.push({
        row: r + 1,
        applicationId: values[r][0],
        company: values[r][1],
        role: values[r][2]
      });
    }
  }

  if (companyMatches.length === 1) {
    const existingRole = normalize_(companyMatches[0].role);
    if (
      existingRole === roleKey ||
      existingRole.includes(roleKey) ||
      roleKey.includes(existingRole)
    ) {
      return companyMatches[0];
    }
  }

  return null;
}


function nextApplicationId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'APP-0001';

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let max = 0;

  ids.forEach(id => {
    const match = String(id).match(/^APP-(\d+)$/i);
    if (match) max = Math.max(max, Number(match[1]));
  });

  return 'APP-' + String(max + 1).padStart(4, '0');
}


function appendUpdate_(ss, email, c, applicationId) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_UPDATES);

  sheet.appendRow([
    email.received,
    applicationId,
    clean_(c.company),
    clean_(c.role),
    normalizeStatus_(c.status, c.classification),
    c.classification,
    Number(c.confidence || 0),
    email.subject,
    email.from,
    email.gmailLink,
    email.messageId,
    email.threadId,
    c.reason || '',
    email.body.substring(0, 500)
  ]);
}


function addReview_(ss, email, c, suggestedAction) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_REVIEW);

  sheet.appendRow([
    email.received,
    c.application_id_hint || '',
    clean_(c.company),
    clean_(c.role),
    c.classification || 'UNCERTAIN',
    Number(c.confidence || 0),
    email.subject,
    email.from,
    email.gmailLink,
    email.messageId,
    email.threadId,
    c.reason || '',
    suggestedAction || ''
  ]);
}


/**
 * Message IDs already handled: Email Updates + Needs Review + Processed.
 * The Processed sheet is what stops ignored emails (alerts, marketing)
 * from being re-sent to the API every run.
 */
function getProcessedMessageIds_(ss) {
  const ids = new Set();

  const sources = [
    { name: CONFIG.SHEET_UPDATES, index: 10 },   // Message ID, column K
    { name: CONFIG.SHEET_REVIEW, index: 9 },     // Message ID, column J
    { name: CONFIG.SHEET_PROCESSED, index: 0 }   // Message ID, column A
  ];

  sources.forEach(src => {
    const sheet = ss.getSheetByName(src.name);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    for (let r = 1; r < values.length; r++) {
      const id = String(values[r][src.index] || '');
      if (id) ids.add(id);
    }
  });

  return ids;
}


/**
 * Records a message ID in the hidden Processed sheet.
 */
function markProcessed_(ss, email) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_PROCESSED);
  if (!sheet) return;
  sheet.appendRow([email.messageId, email.received]);
}


function normalizeStatus_(status, classification) {
  const s = String(status || '').toLowerCase();

  if (classification === 'OFFER' || s.includes('offer')) return 'Offer';
  if (classification === 'REJECTION' || s.includes('reject')) return 'Rejected';
  if (classification === 'INTERVIEW_REQUEST' || s.includes('interview')) return 'Interview';
  if (classification === 'RECRUITER_CONTACT') return 'Recruiter Contact';
  if (s.includes('under review') || s.includes('review')) return 'Under Review';
  if (s.includes('received')) return 'Application Received';
  if (classification === 'APPLICATION_CONFIRMATION') return 'Application Received';
  return 'Applied';
}


function ensureSheets_(ss) {
  [
    CONFIG.SHEET_APPLICATIONS,
    CONFIG.SHEET_UPDATES,
    CONFIG.SHEET_REVIEW,
    CONFIG.SHEET_PROCESSED
  ].forEach(name => {
    if (!ss.getSheetByName(name)) {
      const sheet = ss.insertSheet(name);
      if (name === CONFIG.SHEET_PROCESSED) sheet.hideSheet();
    }
  });
}


function ensureHeaders_(ss) {
  const app = ss.getSheetByName(CONFIG.SHEET_APPLICATIONS);
  if (app.getLastRow() === 0) {
    app.appendRow([
      "Application ID", "Company", "Role", "Application Date", "Status",
      "Last Update", "Last Update Date", "Source Email", "Email Subject",
      "Gmail Link", "Application Key", "Notes", "Next Action"
    ]);
  }

  const updates = ss.getSheetByName(CONFIG.SHEET_UPDATES);
  if (updates.getLastRow() === 0) {
    updates.appendRow([
      "Received", "Application ID", "Company", "Role", "Detected Status",
      "Email Type", "Confidence", "Email Subject", "Sender", "Gmail Link",
      "Message ID", "Thread ID", "AI Reason", "Snippet"
    ]);
  }

  const review = ss.getSheetByName(CONFIG.SHEET_REVIEW);
  if (review.getLastRow() === 0) {
    review.appendRow([
      "Received", "Proposed Application ID", "Company", "Role", "Classification",
      "Confidence", "Email Subject", "Sender", "Gmail Link",
      "Message ID", "Thread ID", "AI Reason", "Suggested Action"
    ]);
  }

  const processed = ss.getSheetByName(CONFIG.SHEET_PROCESSED);
  if (processed.getLastRow() === 0) {
    processed.appendRow(["Message ID", "Received"]);
  }
}


function makeApplicationKey_(company, role) {
  return normalize_(company) + '|' + normalize_(role);
}


function normalize_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\b(intern|internship|summer|fall|spring|2027|2028|2029)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function clean_(value) {
  return String(value || '').trim();
}


function stripMarkdownFence_(text) {
  return String(text)
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
}
