# Wordsmith: conversational research and Word editing

A local Python/browser showcase using **Microsoft Agent Framework**, live Azure web grounding,
and an isolated **Word MCP server**. Research a public topic, create a cited `.docx` report in OneDrive,
then keep chatting to edit its text, formatting, and tables.
Microsoft Graph retrieves/creates/updates files; Python edits the actual Word content.

## Run on Windows

Python 3.11+ is required; this workspace was developed with Python 3.13.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
Copy-Item .env.example .env
# Configure .env locally. Never paste credentials into chat.
.\.venv\Scripts\python.exe -m uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Open <http://127.0.0.1:8000>. The included synthetic sample and file upload/download work without cloud configuration.
The page and static assets revalidate on refresh, and stylesheet/script URLs share a revision
to prevent older cached assets from breaking an updated chat layout.
AI editing requires your real Azure OpenAI deployment; there is no simulated-AI fallback.
Use **Terminal > Run Task > Run Word editing demo** if the VS Code task has been created.
Do not expose this single-user demo on `0.0.0.0` or deploy it as a production service.

## Conversational multi-agent walkthrough

1. Connect Microsoft 365 and browse into your OneDrive **Test** folder.
2. Enable **public web research**. Ask: "Research the recent inflation rate and its causes."
   The coordinator asks which country if the conversation has not established it. Reply "United States."
3. Research requires a connected OneDrive. Each researched report is automatically saved as a new,
   uniquely named file in the visible destination folder; research never overwrites an existing file.
   The chat shows **Saving your report to OneDrive** while uploading and confirms the saved filename.
   If saving fails, the app explicitly reports it, retains a downloadable draft, and exposes
   **Retry report save to OneDrive**. Check OneDrive first if an upload outcome is uncertain.
4. The transcript and activity indicator show coordinator, researcher, writer, and Word MCP stages.
   The report includes dated findings, causes/caveats, and numbered source references with HTTPS URLs.
5. Continue: "Make all its headings blue," then "Make them 20 points and center aligned."
   Review the formatted candidate and apply/discard, or send a follow-up to replace the unapplied proposal.
   Nothing is implicitly applied: revisions start from the current working document, with the pending
   proposal supplied to the coordinator. You can also select paragraphs and request a table.
   Font-family changes (including Libertinus Serif) are supported. Fonts are not installed/embedded;
   the preview or Word may substitute another font when the named family is unavailable locally.
6. Click **Apply & save** to apply the reviewed candidate and automatically update the same OneDrive file.
   Its cloud version is checked to prevent overwriting concurrent edits. To preserve the original,
   choose **Apply as new file** instead; subsequent edits target that new copy.

The app's current folder is the creation destination, not a folder inferred from a chat message.
Opening a different file clears conversation context. **New chat** clears history but keeps the file.
Save/download an unsaved file before replacing it with another research report.
There is no conversation/message-count limit. Unsupported requests do not create empty blocking proposals.
The editor receives the actual formatting schema and explicit nested field examples
(`formats[i].format.color`, for example). An empty edit plan gets one bounded model recheck
against the original request and schema before being returned; this adds a model call for
no-op/unsupported requests, never applies changes automatically, and cannot guarantee model correctness.
Enter sends; Shift+Enter inserts a newline. Long conversations send the most recent complete messages
(up to 24,000 history characters) to the coordinator, plus the current document to the editor.
Older transcript messages remain visible but may need restating if referenced outside that context window.
Request IDs remain deduplicated across conversation resets in the same session; receipts do not cache
entire transcript/document snapshots per message. Replaying a request returns the current session view
without repeating research, edits, or uploads. State remains in memory, expires after one hour of inactivity,
and disappears on server restart. Per-message/document size and model/service quotas still apply.

### Agent and MCP architecture

```text
Browser chat -> FastAPI session -> Agent Framework Coordinator
                                 |- clarify / answer using conversation history
                                 |- WebResearcher (Azure Responses + required hosted web search)
                                 |    -> ReportWriter (structured, numbered source references)
                                 |    -> Word MCP create_research_report / export_candidate
                                 `- WordEditor (structured text/format/table plan)
                                      -> Word MCP inspect_document / preview_edits / export_candidate

Host-controlled create/update -> Microsoft Graph (delegated Files.ReadWrite)
```

These are separate `agent_framework.Agent` instances orchestrated by Python, not UI labels around
one prompt. Conversation context is stored per browser session and passed to the coordinator;
the editor receives the current snapshot through MCP, not stale prose from earlier turns.
The existing `/api/proposal` endpoint remains for compatibility; the browser now uses `/api/chat`.

The private MCP worker runs via `python -m app.word_mcp`, using the actual MCP stdio protocol.
Each workflow gets a fresh worker with no shared document state, network listener, cloud tools,
arbitrary file-access tools, or cloud tokens. The host—not the model—loads/exports document bytes,
validates selection/version constraints, manages review/apply, and authorizes Graph writes.
The worker exposes `load_document`, `inspect_document`, `preview_edits`, `create_research_report`,
and `export_candidate`. It stages candidates only; it cannot apply or save to OneDrive.

The workflow uses the official MCP SDK client directly with Agent Framework orchestration.
The installed Agent Framework high-level MCP adapter was tested and failed its initialization against
MCP SDK 2.x (`protocolVersion` versus `protocol_version`); the direct SDK integration avoids patching
third-party internals. MCP is pinned to the tested 2.2.0 version.

### Live research setup and boundaries

Your existing Azure model endpoint/deployment and **Entra authentication** are reused. The deployment
must support structured chat outputs plus the **Responses API and hosted `web_search` tool**.
No new Foundry project, separate Bing key/connection, or hosted-agent deployment is required.
See [Microsoft Agent Framework web search](https://learn.microsoft.com/agent-framework/agents/tools/web-search)
and [Azure web search](https://learn.microsoft.com/azure/ai-foundry/agents/how-to/tools/web-search?view=foundry).

Web search is preview, billable, and subject to availability/model/subscription policy. An unsupported
deployment, blocked tool, expired sign-in, or missing citations produces an explicit error—never a fake
research result or model-memory fallback. For errors, confirm model/tool support and ask the Azure
administrator whether `OpenAI.BlockedTools.web_search` is enabled. No permissions/resources are
created automatically.

Public search queries/results can flow outside Azure geographic/compliance boundaries; the Data
Protection Addendum does not apply to Bing grounding data.
Review [Bing grounding terms](https://www.microsoft.com/bing/apis/grounding-legal-enterprise)
and [pricing](https://www.microsoft.com/bing/apis/grounding-pricing).
Do not include confidential information or secrets in research requests. The researcher receives only
the public research task and date, not the loaded Word document or full conversation.
Returned pages are untrusted evidence. Only provider-returned, valid HTTPS citation URLs become report
sources; each section must cite known source IDs. This validates citation provenance and structure,
not the truth of every generated claim: review the official releases before relying on the report.

Research requires connected Microsoft 365 and public-search consent, with a visible destination folder.
Report creation in OneDrive is automatic. Request IDs prevent a repeated chat request from
creating duplicate files in the same live session. Unknown save outcomes leave the local draft
available and display an error; check OneDrive before requesting another creation.

## Azure OpenAI configuration

If you do not already have an approved resource, use [Azure portal](https://portal.azure.com) to create an
**Azure OpenAI** resource in your subscription, resource group, and supported region. Resource/model usage can incur Azure charges.
Open its Microsoft Foundry portal and deploy a structured-output-capable chat model, for example a compatible GPT-4.1 deployment.
Give the deployment an easy-to-recognize name such as `word-demo`. No Azure hosting for this app is needed.

Set these in `.env`, then restart:

```dotenv
AZURE_OPENAI_ENDPOINT=https://YOUR-RESOURCE.openai.azure.com
AZURE_OPENAI_DEPLOYMENT=your-deployment-name
AZURE_OPENAI_AUTH=entra
```

| Setting | Where to get it |
|---|---|
| `AZURE_OPENAI_ENDPOINT` | Azure portal > your Azure OpenAI resource > **Keys and Endpoint** > endpoint. Use the resource URL, not a portal page URL. |
| `AZURE_OPENAI_DEPLOYMENT` | Microsoft Foundry > the resource's **Deployments / Models + endpoints** > your deployment's exact name, such as `word-demo`. |
| `AZURE_OPENAI_AUTH` | Set `entra` for local Azure sign-in, or explicitly set `key` for API-key authentication. |
| `AZURE_OPENAI_API_KEY` | Only for `key` mode: copy a key from the resource's **Keys and Endpoint** into your local `.env`. Never share it in chat. |

The endpoint may also end in `/openai/v1/`. This demo supports public Azure resource endpoints, not sovereign clouds.
For a Foundry resource, the resource OpenAI base URL can use its `.services.ai.azure.com` host, for example
`https://YOUR-RESOURCE.services.ai.azure.com/openai/v1/`.
Do not include an operation suffix such as `/responses` or `/chat/completions`: the SDK appends the operation itself.
The Foundry **project endpoint** (`https://YOUR-RESOURCE.services.ai.azure.com/api/projects/YOUR-PROJECT`)
is a different API and is not accepted by this app. Use the resource's OpenAI endpoint shown in the deployment's
OpenAI SDK sample, not the project endpoint. This demo makes Chat Completions requests, not Responses requests.
Use a deployment supporting **Chat Completions structured outputs**, such as a compatible GPT-4o or GPT-4.1 deployment.
The deployment name, not merely the underlying model name, is passed as `model`.

For passwordless local use, sign in with your Azure developer tooling (for example `az login` if Azure CLI is installed).
Your identity needs model inference permission, such as **Cognitive Services OpenAI User**, on the resource.
The app uses local `DefaultAzureCredential` with managed-identity probing disabled and scope `https://ai.azure.com/.default`.

Alternatively, explicitly select key authentication for an existing resource:

```dotenv
AZURE_OPENAI_AUTH=key
AZURE_OPENAI_API_KEY=your-key-stored-only-in-this-local-file
```

**If API-key authentication is disabled on your resource, do not enable it for this demo.**
Keep `AZURE_OPENAI_AUTH=entra`, leave `AZURE_OPENAI_API_KEY` empty, and use your authorized local Azure identity.
No Entra client secret is needed for either the local Azure identity or this app's OneDrive sign-in.

Do not commit `.env`. Credentials and model tokens stay on the Python server, not in browser JavaScript.
When **Send to agent team** is clicked, the coordinator receives the message and bounded conversation
history; the editor receives supported editable text and basic formatting/structure metadata from the
current document. These calls use the configured Azure endpoint. The web researcher receives only
the public research task/date after research is enabled, not the document snapshot.
Use synthetic documents until that service is approved for the intended data. This app is local, **not offline**.

## Microsoft 365 / OneDrive configuration

1. Use a single-tenant **Microsoft Entra app registration** in your work/school tenant.
   In [Entra admin center](https://entra.microsoft.com), go to **App registrations > New registration**,
   name it `Word Editing Demo`, choose **Accounts in this organizational directory only**, and register.
   On **Overview**, copy **Directory (tenant) ID** to `ENTRA_TENANT_ID` and **Application (client) ID** to `ENTRA_CLIENT_ID`.
   Do not use the Object ID.
2. Under **Authentication > Add a platform > Mobile and desktop applications**, add `http://localhost`.
   Leave the suggested `nativeclient`, LiveSDK, and `msal...://auth` checkboxes unchecked for this demo.
   Use **Add URI** / the custom redirect URI field to enter `http://localhost`, then save/configure it.
   Those suggested redirects belong to other authentication flows; this app uses a system-browser loopback redirect.
   MSAL opens a system browser and uses a temporary loopback listener. No client secret is needed.
   This redirect is not the app's `http://127.0.0.1:8000` address; do not configure this registration as a SPA or confidential web app.
3. Add Microsoft Graph **delegated** `Files.ReadWrite`; grant the consent your tenant requires.
   Use **API permissions > Add a permission > Microsoft Graph > Delegated permissions**.
   The **+ Add a permission** button is directly above the configured-permissions table.
   In the permission picker, search for `Files.ReadWrite`, expand **Files**, select exactly `Files.ReadWrite`,
   and click **Add permissions**. `User.Read` alone is not sufficient.
   You may leave an existing `User.Read` entry; the app does not need it for file operations.
   If **Grant admin consent** is disabled, an authorized tenant administrator must grant consent when your policy requires it.
   Do not add application permissions or `Files.ReadWrite.All` for this demo.
4. Configure:

   ```dotenv
   ENTRA_TENANT_ID=your-directory-guid
   ENTRA_CLIENT_ID=your-application-guid
   ```

5. Restart the server and click **Connect Microsoft 365**.
   To restart the existing demo, press **Ctrl+C** in its VS Code task terminal, then rerun the task or the launch command above.

OneDrive sign-in and Azure OpenAI authentication are separate. A Graph access token does not grant model inference access.
Tenant consent policy, Conditional Access, resource firewall rules, and account permissions still apply.
Sign-in has a three-minute timeout. Work/school OneDrive is the intended live flow.
During sign-in, a visible **Open Microsoft sign-in** link opens the local welcome page if the default
browser window is hidden or fails to open. Click **Continue to Microsoft sign-in**, complete account selection/consent,
then return to the demo. This must be done on the same computer as the Python server.
Being signed into OneDrive in a browser is not the same as authorizing this app.
Shared shortcuts, SharePoint library discovery, and personal-account registration are not included.

### Work with your Test folder and update the same file

1. In OneDrive in your browser, open **Test > New > Word document**.
2. Name the document, add a few paragraphs (the demo cannot edit an empty document), wait for Word to save, and close its editor.
3. In this demo, click **Connect Microsoft 365**, finish sign-in with the same account, and return here.
4. Open **Test** in the demo's OneDrive list. Use **Refresh folder** if the new document is not listed, then click its filename.
5. Enter an AI instruction, generate a proposal, and review it.
6. Click **Apply & save**. This automatically updates that document, not a sibling copy, without a second confirmation.
7. Open the saved document link or refresh OneDrive to see the changes.

The update sends the opened file's cloud eTag as `If-Match`. A cloud-version conflict blocks the update;
download your working copy before reopening the cloud file. Close Word before using this demo to avoid concurrent editing.
If you want to preserve the original instead, choose **Apply as new file** before applying.
**Save as new file** in the document toolbar saves an additional copy of the current working document.
If an automatic save fails, the applied candidate stays available locally with an explicit error and
**Retry save to same file**. Download it or check OneDrive before retrying an uncertain upload.

## Showcase walkthrough

1. Click **Try the sample proposal**. Inspect the heading, summary, payment terms, and table text.
2. Click **30 days -> 45 days**, then **Send to agent team**.
3. Show the **before/after** difference. The working document and OneDrive are still unchanged.
4. Click **Apply locally** for the sample, then **Download .docx**. Open the download in Word:
   the bold payment term, italic budget, table, image, header, and footer remain in the file.
5. Try **Make headings blue**. Review the colored **Proposed result** below the unchanged working preview,
   then apply. Try "Make the headings 20 points, bold, and center aligned" too.
6. Try **Scope list -> table**. The proposal shows a genuine table with a header and two scope rows.
   Apply and download: the exported file contains a real Word table, not a picture or pipe-separated text.
   For other content, select consecutive paragraphs/list items with their checkboxes and use **Selection -> table**.
   Selection limits all AI operations to those blocks; leave it empty for a whole-document instruction.
7. To demonstrate Graph, upload `demo\sample.docx` to your OneDrive **Test** folder using OneDrive itself.
   Connect Microsoft 365, open that file in the demo, repeat an edit, and click **Apply & save**.
8. Open the document in Word. Your same file now has the reviewed text, colors, and tables.
   Choose **Apply as new file** instead if you want to keep the original unchanged.

The sample is intentionally fictional. Proposing and reviewing never write to OneDrive.
Applying a local sample/upload stays local; applying a OneDrive document automatically saves it.
Research writes a new file only when the explicit per-message creation option is enabled.
You can discard a proposal or clear the working document. Download/save before clearing if you want to keep edits.
**Refresh state** recovers the current document, conversation, and pending proposal after a connection problem.

## Editing guarantees and limits

- Python edits the existing OOXML package, not an HTML-to-Word reconstruction.
- Unmodified ZIP parts are copied byte-for-byte, including images, styles, headers, and footers.
- Main-body paragraphs, headings, lists, and ordinary table-cell text are supported.
- Text-only edits retain existing runs/properties; inserted text inherits formatting at its insertion point.
  A rewrite can change text wrapping and pagination: check the resulting layout in Word.
- AI can format whole editable paragraphs/headings/cell paragraphs: RGB font color, 6-72pt font size
  (half-point increments), bold/italic, and left/center/right/justified alignment.
  Only specified properties change. Formatting applies directly to existing text, not global Word style definitions.
- Consecutive editable main-body paragraphs/list items can become a **real Word table**:
  1-8 columns, up to 50 data rows, optional new bold header, explicit table/grid/cell widths, borders, and padding.
  This intentionally rebuilds the selected layout and removes list markers. Cell text uses the first source
  paragraph's base formatting, rather than preserving mixed inline emphasis in the converted source.
  Source text and order must be preserved; only whitespace, pipe and semicolon separators may be omitted.
  Rewriting/reordering the content requires a separate reviewed text edit before conversion.
- The browser shows basic inline colors, sizes, bold/italic, alignment, and actual tables, including the
  **proposed result before apply**. It is not exact Word rendering: pagination, theme effects,
  merged/nested table layout, fonts, list numbering, headers, footers, and images are not fully rendered.
  The document canvas stays paper-colored so Word ink remains visible in dark mode.
- Native Track Changes, arbitrary document creation, table merging/resizing, and collaborative editing are not supported.
- Complex regions (fields, hyperlinks, comments/bookmarks, content controls, revisions, text boxes, tabs/line breaks)
  are read-only and excluded from AI proposals. Existing revisions are preserved; deleted text is not previewed.
- Protected, signed, encrypted, macro-enabled, malformed, and unsupported Word packages are rejected.
- Limits: **10 MiB compressed**, **50 MiB expanded**, **2,000 ZIP parts**, **30,000 editable text characters**,
  **2,000 instruction characters**, and **50 operations per proposal**. Input is not silently truncated.
- Proposals must match the exact document snapshot, target known editable blocks, match source text uniquely,
  and contain no overlapping operations (text plus formatting on the same block is allowed).
  Table targets cannot overlap any other operation. Apply installs the exact validated document shown in review,
  is atomic, and rejects stale versions.
- **Save as new file** and **Apply as new file** use a new name and an upload session with conflict behavior `fail`.
- **Apply & save** updates the opened item's content with an `If-Match` version precondition.
  Applying is the save action; there is no second overwrite confirmation. Generating and reviewing never update OneDrive.
  A failed update keeps your applied in-memory working copy and reports that saving was not confirmed.
  If a save times out, check OneDrive before retrying: completion might be unknown.
- Files/tokens/proposals live in server memory, not a document database. Reset clears document state; disconnect also clears the token cache.
  Sessions expire after an hour of inactivity, with at most four browser sessions. Restarting clears all state.
- The server binds to loopback, restricts hosts/origins, uses an HTTP-only same-site cookie, and does not log document text or tokens.
  This is not a production security boundary against other processes on the same machine.

## Validation and fixture generation

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m pyright --pythonpath .\.venv\Scripts\python.exe
node --check app\static\app.js
```

Tests use synthetic documents and explicit test doubles for Azure/Graph; they do not prove live tenant/model readiness.
Complete the OneDrive walkthrough above with your configured services to validate those integrations.

The sample is already included. Node is needed only to regenerate it:

```powershell
npm install
npm run sample
```

## Troubleshooting

| Symptom | Action |
|---|---|
| Setup needed | Fill `.env` locally and restart; placeholders are not valid configuration. |
| Microsoft sign-in cannot complete | Check tenant/client GUIDs, desktop `http://localhost` redirect, consent, and Conditional Access. |
| Connect appears to do nothing | Refresh the page after restarting the server. Click Connect, then use the visible **Open Microsoft sign-in** link if a separate window is not visible. Complete the flow within three minutes. |
| OneDrive access denied | Verify delegated `Files.ReadWrite` and permission to the specific source folder. |
| Azure authentication failed | Check your local Azure identity or explicitly selected API key. |
| Azure denied model access | Check resource data-plane RBAC, network/firewall access, and deployment. |
| Azure rejected request | Confirm Chat Completions structured-output support and deployment name; narrow the instruction. |
| Proposal cannot match text | Ask for a more specific change; read-only and ambiguous regions cannot be edited safely. |
| File is still encrypted or rights-managed | A downloaded `.docx` can remain an encrypted Office container even after a visible label changes. This Python demo cannot open protected content. Use an organization-approved synthetic, unprotected demo file, or ask the owner/admin to confirm protection status and that the changed version has saved to OneDrive. Do not remove required protection to use this demo. |
| DOCX ZIP package is unreadable or damaged | Download a fresh cloud copy and retry. This is a package-integrity error, not proof of encryption. |
| Save completion unknown | Inspect the source OneDrive folder before retrying. Never assume a timeout means nothing was saved. |
| No usable files on a folder page | Use Next page or choose another folder; only owned folders and `.docx` files are shown. |
| Local connection failed | Keep the Python server running; return to the same loopback URL and use Refresh state. |
