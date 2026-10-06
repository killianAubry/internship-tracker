# Internship Application Tracker

A Google Apps Script that scans your Gmail for internship application emails, classifies them with an LLM on Groq's free tier, and keeps a Google Sheet up to date with every application and its status.

## How it works

```
Gmail keyword filter -> Groq LLM classifier -> Application matching -> Google Sheet
```

1. A daily trigger searches Gmail for likely application emails (confirmations, interview requests, rejections, offers).
2. Each new email is sent to Groq (`openai/gpt-oss-120b` by default), which returns strict JSON: classification, confidence, company, role, and status.
3. The script matches the email to an existing application (by `APP-####` hint, then company + role) or creates a new one.
4. Low-confidence or ambiguous emails go to a **Needs Review** sheet instead of changing your data.

## Sheets

| Sheet | Purpose |
|---|---|
| `Applications` | One row per application with current status |
| `Email Updates` | Audit log of every accepted email |
| `Needs Review` | Ambiguous emails and API failures to check manually |
| `Processed` (hidden) | Message IDs already classified, so nothing is re-sent to the API |

## Setup

1. Create a Google Sheet, then open **Extensions -> Apps Script**.
2. Paste in `src/Code.gs`.
3. Get a free key at [console.groq.com/keys](https://console.groq.com/keys) and add it under **Project Settings -> Script properties** as `GROQ_API_KEY`.
4. Run `setupTracker()` once and authorize access. This creates the sheets and a daily trigger.
5. Optional: run `scanInternshipEmails()` manually to process the first batch.

Or deploy with [clasp](https://github.com/google/clasp): `clasp create --type sheets`, then `clasp push` (set `rootDir` to `src` in `.clasp.json`).

## Free-tier behavior

- Rate limits (HTTP 429) are retried using the `retry-after` header. If the limit persists, the run stops and the remaining emails are picked up next run.
- Auth or model errors (401/403/404) also stop the run instead of flooding Needs Review.
- Runs stop at about 4.5 minutes to stay under the Apps Script 6-minute limit.
- If a run fails or a model changes, run `retryFailedLlmRows()` to reclassify rows marked as API failures.

## Configuration

Edit the `CONFIG` object at the top of `src/Code.gs`:

- `GROQ_MODEL`: `openai/gpt-oss-120b` (accuracy) or `openai/gpt-oss-20b` (lighter on limits). Check [Groq's model list](https://console.groq.com/docs/models) if you get a 404.
- `GMAIL_QUERY`: which emails are candidates.
- `AUTO_ACCEPT_CONFIDENCE`: below this, emails go to Needs Review.
- `MAX_BODY_CHARS`, `CALL_SPACING_MS`, `DAILY_TRIGGER_HOUR`.

## Privacy

Email subjects, senders, and the first 3,000 characters of each candidate email are sent to Groq. Groq's free tier has no privacy guarantees, so review their data policy first.
