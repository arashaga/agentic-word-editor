# Wordsmith Backend Walkthrough

This guide explains how the backend loads, edits, reviews, and saves Word documents. It is
written for a software engineer who is new to this codebase. The main idea is:

> **The AI suggests changes. Python checks and applies them. Microsoft Graph moves files
> to and from OneDrive.**

The browser never talks directly to Microsoft Graph. It calls this app's FastAPI endpoints,
and the backend controls the document and cloud operations.

## The big picture

```text
Browser
  |  HTTP requests to /api/...
  v
FastAPI (app/main.py)
  |-- keeps this browser's document, proposal, chat, and OneDrive connection in a Session
  |
  |-- WordDocument (app/documents.py)
  |     reads and edits the DOCX package on the backend
  |
  |-- AgentTeam / AzureEditor
  |     asks the model for a structured edit plan
  |
  |-- isolated Word MCP worker (app/word_mcp.py)
  |     validates and stages a candidate document; it cannot save to OneDrive
  |
  `-- OneDrive (app/onedrive.py)
        signs in through MSAL and uses Microsoft Graph to browse/download/upload files
```

### A few important words

- **DOCX**: a Word file is a ZIP package containing XML and other parts such as images.
- **Block**: a paragraph or other text unit that this app can show and potentially edit.
- **Edit plan**: structured instructions such as “replace this exact text” or “make this
  paragraph blue.” It is not yet a changed document.
- **Candidate**: a copy of the document with a proposed plan applied. The user can preview
  it before accepting it.
- **ETag**: a version identifier Microsoft Graph returns for a cloud file. The app uses it
  to detect concurrent changes.
- **Delegated permission**: Graph acts as the signed-in user and is limited by that user's
  access and the consented scope.

## Who owns what?

| File | Responsibility |
|---|---|
| [`app/main.py`](app/main.py) | FastAPI routes, browser sessions, request validation, review/apply flow, and deciding when cloud saves happen. |
| [`app/documents.py`](app/documents.py) | Reads, previews, validates, and edits DOCX XML packages. |
| [`app/onedrive.py`](app/onedrive.py) | Microsoft sign-in, Graph requests, downloads, safe updates, and new-file uploads. |
| [`app/config.py`](app/config.py) | Reads local environment configuration and checks required settings. |
| [`app/ai.py`](app/ai.py) | Gets a structured editing plan from Azure OpenAI. |
| [`app/agents.py`](app/agents.py) | Routes chat to research/edit/answer flows and coordinates the model and MCP worker. |
| [`app/word_mcp.py`](app/word_mcp.py) | Provides isolated Word tools that inspect a snapshot and stage/export a candidate. |
| [`app/errors.py`](app/errors.py) | `DemoError` carries a safe message and HTTP status for expected failures. |

Most requests enter through `create_app()` in [`app/main.py`](app/main.py#L186). Routes use
the `Session` associated with a browser cookie. `SessionStore` keeps sessions in memory, and
each session owns its own `OneDrive` object and current document. This is demo-style
single-process storage: the sessions are not a durable database and do not survive a server
restart. See [`Session`](app/main.py#L38) and [`SessionStore`](app/main.py#L64).

## 1. Microsoft sign-in: how the backend gets permission

The Graph configuration is `ENTRA_TENANT_ID` and `ENTRA_CLIENT_ID`, loaded in
[`app/config.py`](app/config.py). No client secret is used for this desktop-style sign-in.

1. The browser calls `POST /api/auth/login` ([route](app/main.py#L263)).
2. The route calls `session.drive.sign_in()` in [`app/onedrive.py`](app/onedrive.py#L110).
3. MSAL creates a `PublicClientApplication` for the configured tenant and opens an
   interactive Microsoft sign-in using a temporary local loopback port.
4. The app requests the delegated Graph scope `Files.ReadWrite`
   ([scope constant](app/onedrive.py#L19)).
5. After sign-in, the MSAL client is held by that browser's in-memory `OneDrive` object.
   Later Graph calls ask MSAL for a token using `acquire_token_silent()` in
   [`_token()`](app/onedrive.py#L167).
6. Each Graph request sends that token as `Authorization: Bearer ...`.

The login route and `GET /api/auth/pending` support the browser's sign-in UX. Logging out
removes the account from the MSAL client and clears the session's open document
([logout route](app/main.py#L273)).

**Keep the two kinds of authentication separate:**

- **Microsoft Graph / OneDrive** uses MSAL with the user's Microsoft 365 sign-in and the
  `Files.ReadWrite` delegated scope.
- **Azure OpenAI / agents** uses the Azure settings and credentials in `app/config.py`,
  `app/ai.py`, and `app/agents.py`. Those credentials are not the Graph access token.

## 2. Opening a Word file from OneDrive

The opening flow is:

1. `GET /api/drive` calls `OneDrive.browse()` to list folders and ordinary `.docx` files
   ([route](app/main.py#L280), [Graph implementation](app/onedrive.py#L194)).
2. When a file is selected, the browser sends its item ID to `POST /api/document/onedrive`
   ([route](app/main.py#L288)).
3. `OneDrive.download()` asks Graph for the file metadata, checks that the item is a
   supported Word file, then downloads its content ([implementation](app/onedrive.py#L216)).
4. The backend builds a `SourceReference` containing the Graph item ID, parent folder,
   filename, web URL, and ETag. This is the information needed to update the same cloud
   file safely later.
5. `_load()` creates a `WordDocument` from the downloaded bytes and stores the document
   and source reference in the session ([helper](app/main.py#L167)).

Graph may redirect the content download to a temporary file URL. The app validates those
URLs before requesting them, does not automatically follow redirects, and only sends the
Graph bearer token to the Graph host. The download is size-limited. If Graph supplied an
ETag, the app checks it again after download; if the file changed during that operation,
opening fails rather than silently using an uncertain snapshot.

## 3. How a DOCX becomes editable text

[`WordDocument`](app/documents.py#L187) does not launch Microsoft Word. It works directly
with the DOCX package:

1. It checks the filename, file size, ZIP structure, and required Word package parts.
2. It reads `word/document.xml`, which contains the document body, paragraphs, runs, tables,
   and formatting properties.
3. It turns readable paragraphs into `Block` objects with IDs such as `p3`, text, kind
   (heading, paragraph, list, or table), formatting, and an editable/read-only flag.
4. It keeps the original package parts so images, headers, and other content can be
   preserved when it rebuilds the file.

Some Word content is deliberately read-only. For example, text boxes, content controls,
tracked revisions, fields, and other complex XML are not treated as ordinary editable
paragraphs. `ai_blocks()` sends only supported editable blocks to the model and applies a
text limit ([method](app/documents.py#L368)). The preview is a content preview, not a
pixel-perfect rendering of Word pagination or page layout.

For a text replacement, `preview()` checks that the targeted block exists, is editable,
and contains the exact original text exactly once. It rejects ambiguous, overlapping, or
unsafe replacements ([method](app/documents.py#L376)). Formatting and table operations
have additional schema and content checks in `preview_plan()`
([method](app/documents.py#L407)).

When applying the plan, `apply_plan()` starts from a fresh `WordDocument` built from the
current bytes, changes the XML, writes a new ZIP package, and parses the result again as a
final validation step ([method](app/documents.py#L580)). This creates a new document
object; it does not overwrite the original object in place.

## 4. Editing is a review-first pipeline

For chat-based editing, the main route is `POST /api/chat`
([route](app/main.py#L392)). At a high level:

1. FastAPI checks the request's document version, selected block IDs, and session state.
2. `AgentTeam.run()` asks the Coordinator what the user wants: edit, research, answer, or
   clarification ([workflow](app/agents.py#L173)).
3. For an edit, the host starts a fresh Word MCP process over standard input/output. It
   sends the current DOCX snapshot to the worker, then asks it to inspect the document.
4. The WordEditor model returns an `EditPlan`: structured operations, not a file and not
   executable code. The plan has text edits, format edits, and/or table conversions.
5. The MCP worker validates the plan and stages a candidate. The host exports that
   candidate and validates it again as a `WordDocument`.
6. FastAPI stores the plan, candidate, and document version in a `Proposal`. The user sees
   the preview and chooses to apply or discard it.

The MCP tools live in [`app/word_mcp.py`](app/word_mcp.py). Their boundary is important:
the worker can load a snapshot, inspect it, validate edits, and export a candidate. It
**cannot apply changes to the session itself or call Microsoft Graph**. The host controls
the final apply and any cloud write. The worker starts with a minimal environment rather
than inheriting Azure or Graph tokens; see `word_tools()` in
[`app/agents.py`](app/agents.py#L114).

There is also a compatibility endpoint, `POST /api/proposal`, which calls
`AzureEditor.propose()` directly ([route](app/main.py#L314),
[editor](app/ai.py#L68)). It follows the same basic idea: the model proposes a structured
plan, and the backend validates it before creating a preview.

### Apply or discard

- `POST /api/proposal/discard` drops the pending proposal; the working document stays
  unchanged ([route](app/main.py#L379)).
- `POST /api/proposal/apply` checks that the proposal ID and document version still match,
  then makes the reviewed candidate the session's working document
  ([route](app/main.py#L341)).
- If the document came from OneDrive, apply normally attempts to update that same cloud
  file. The user can instead choose `save_as_new` to keep the original unchanged.
- If the cloud save fails, the local applied document is retained and the error is shown.
  The backend does not pretend an unconfirmed save succeeded.

The document version is a SHA-256 hash of its bytes. Checking it prevents a stale browser
preview or proposal from being applied after the current working document has changed.

## 5. Updating the same OneDrive file

The same-file update is deliberately conditional:

1. The session remembers the ETag returned when the file was opened.
2. `OneDrive.update()` sends the new DOCX bytes to
   `PUT /me/drive/items/{item-id}/content`.
3. The request includes `If-Match: <saved ETag>` ([implementation](app/onedrive.py#L263)).
4. Microsoft Graph accepts the write only if the cloud file is still at that version.
   If someone else changed it, Graph returns HTTP 412; the app reports a conflict rather
   than overwriting their newer file.
5. On success, the app saves the new ETag and filename/URL back into the session's
   `SourceReference` ([helper](app/main.py#L175)).

The explicit retry/update endpoint is `POST /api/document/update`
([route](app/main.py#L513)). It only runs when a OneDrive source exists and the working
document has unapplied-to-cloud changes.

## 6. Saving a separate copy or creating a report

### “Save as new”

`OneDrive.save_copy()` makes a unique edited filename and calls `_create_upload()`
([copy method](app/onedrive.py#L289)). `_create_upload()`:

1. Asks Graph to create an upload session in the source file's parent folder.
2. Sets conflict behavior to `fail`, so it will not silently replace a same-named file.
3. Uploads the bytes to the returned upload URL and checks Graph's final response.

The upload URL is a temporary capability URL. The app validates its host and uploads the
file without forwarding the Graph bearer token to that URL. The same upload helper is
used to create a new report in a selected folder
([new document method](app/onedrive.py#L298)).

### Research reports

Research is a separate workflow from editing. It requires the user to be connected to
OneDrive. The agents research public sources, compose a structured report, and use the Word
MCP worker to create a real DOCX. Then the host calls `OneDrive.create_document()` to save
it automatically ([chat route](app/main.py#L392)). If that save fails, the draft remains
available in the session for download; the UI is told that saving was not confirmed.

## 7. A few safety and reliability controls

- **Same-origin demo API checks:** non-GET API requests must have the demo request header
  and must not be cross-site. The session cookie is HTTP-only and `SameSite=Strict`.
  See the middleware in [`app/main.py`](app/main.py#L211).
- **Input and document limits:** request fields, file bytes, expanded ZIP size, number of
  package parts, and model input length are bounded.
- **Safe XML parsing:** XML external entities and network access are disabled in
  [`_xml()`](app/documents.py#L147).
- **Structured changes only:** model output must match Pydantic schemas and pass the same
  document validation used for preview. The model never supplies raw XML or Python code
  to execute.
- **Cloud URL validation:** Graph and temporary content URLs are checked before requests
  are made ([`_cloud_url()`](app/onedrive.py#L58)).
- **Expected failures are explicit:** Graph status codes become understandable
  `DemoError` messages in [`_check()`](app/onedrive.py#L82), and FastAPI maps those to
  JSON errors in [`main.py`](app/main.py#L196).
- **No production claim:** this application keeps session state in process memory and is
  designed as a local showcase, not a production multi-user service.

## 8. Suggested code-reading path

To follow an edit request in the source, read these in order:

1. [`app/main.py`](app/main.py#L392) — `/api/chat` validates session state and receives the
   agent result.
2. [`app/agents.py`](app/agents.py#L173) — `AgentTeam.run()` routes the request and
   coordinates the editor and MCP worker.
3. [`app/word_mcp.py`](app/word_mcp.py#L14) — the worker's tools load, validate, and export
   a candidate.
4. [`app/documents.py`](app/documents.py#L187) — `WordDocument` parses and changes the
   actual DOCX package.
5. [`app/main.py`](app/main.py#L341) — `/api/proposal/apply` commits the reviewed
   candidate to the session and requests the cloud save.
6. [`app/onedrive.py`](app/onedrive.py#L263) — Graph conditionally updates the source file.

To follow OneDrive authentication and opening, start at
[`app/main.py`](app/main.py#L263), then follow
[`OneDrive.sign_in()`](app/onedrive.py#L110),
[`OneDrive._token()`](app/onedrive.py#L167),
[`OneDrive.download()`](app/onedrive.py#L216), and finally `_load()`.
