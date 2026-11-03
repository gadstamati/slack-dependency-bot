# Slack Dependency & Initiatives Audit Bot

Automated documentation and dependency auditing tool developed in **Google Apps Script (V8 runtime)** for PedidosYa. It scans initiative tracking sheets to identify missing documentation (PRD/PPT links and Definition of Done) and notifies responsible Product Managers via direct Slack messages.

---

## 🔒 Security & Slack API Scopes Justification

This application strictly adheres to the **Principio de Mínimo Privilegio (Principle of Least Privilege)**. It only requests the exact scopes required for its operation:

| Scope | API Method Called | Justification |
| :--- | :--- | :--- |
| **`users:read.email`** | `GET https://slack.com/api/users.lookupByEmail` | Resolves the recipient's Slack User ID using their corporate email address (`@pedidosya.com`) mapped from the tracking sheet. |
| **`chat:write`** | `POST https://slack.com/api/chat.postMessage` | Sends the structured audit summary notification directly to the responsible Product Manager via Direct Message. |

> **Note on Scope Assessment:**  
> An earlier version of the bot manifest requested 6 scopes. The permission set has been narrowed down to **only 2 scopes** (`users:read.email` and `chat:write`). No other Slack scopes or administrative permissions are needed or used.

---

## 🛡️ Credential & Secrets Management

- **No Hardcoded Secrets:** The codebase contains **zero** hardcoded Slack tokens, secrets, or API keys.
- **Secure Key Storage:** The Slack Bot User OAuth Token is loaded at runtime via Google Apps Script's encrypted property store:
  ```javascript
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  ```
- **Rate Limiting & Memory Caching:** The script caches email-to-Slack-ID lookups in-memory during the execution lifecycle (`emailToSlackIdCache`) to minimize API round-trips, and enforces rate-limiting pauses (`Utilities.sleep(600)`) between messages to prevent Slack API abuse.
- **Dry-Run Mode:** Built-in simulation flag (`CONFIG.DRY_RUN: true`) for development and auditing verification without dispatching live messages.

---

## 🏗️ Architecture & Code References

The entire application logic is contained in `Code.gs`:

- **Audit Engine (`runDocumentationAudit`)**: Scans domain tabs, validates PRD/PPT links and DoD completeness, and groups action items by PM.
- **Slack Dispatcher (`sendSlackDMByEmail`)**: Handles the two Slack Web API calls (`users.lookupByEmail` and `chat.postMessage`).
- **Data Resolution (`loadSquadToPmMap`, `resolveSmartChipsNatively`)**: Resolves PM assignments and Google Drive/Docs Smart Chips natively within Google Sheets without external Cloud APIs.
- **Scheduler (`installWeeklyTriggers`)**: Automates execution on scheduled business days (Mondays and Thursdays at 9:00 AM ART).

---

## 📦 Files

- `Code.gs`: Core Apps Script source code.
- `appsscript.json`: Google Apps Script manifest and OAuth scopes configuration.
- `README.md`: Architectural documentation and security assessment reference.
