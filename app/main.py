"""Serve the local browser demo and own session, review/apply, and OneDrive save state."""

from dataclasses import dataclass, field
import asyncio
import logging
from pathlib import Path
import threading
import time
from typing import Annotated
from urllib.parse import quote
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from starlette.middleware.trustedhost import TrustedHostMiddleware

from app.ai import AzureEditor
from app.agents import AgentTeam, ChatTurn
from app.config import ROOT, Settings
from app.documents import DOCX_TYPE, MAX_FILE_BYTES, Change, EditPlan, WordDocument
from app.errors import DemoError
from app.onedrive import GraphItem, OneDrive, SourceReference

logger = logging.getLogger("word_showcase")
COOKIE = "word_demo_session"


@dataclass
class Proposal:
    """Review metadata and the exact validated candidate for one document version."""

    id: str
    version: str
    plan: EditPlan
    changes: list[Change]
    candidate: WordDocument | None


@dataclass
class Session:
    """In-memory document, delegated drive client, transcript, and request receipts."""

    id: str
    drive: OneDrive
    touched: float = field(default_factory=time.monotonic)
    lock: threading.RLock = field(default_factory=threading.RLock)
    document: WordDocument | None = None
    source: SourceReference | None = None
    proposal: Proposal | None = None
    modified: bool = False
    chat: list[ChatTurn] = field(default_factory=list)
    activity: str = "Ready"
    chat_results: dict[str, dict[str, object]] = field(default_factory=dict)
    research_draft: bool = False
    save_error: str | None = None

    def clear_document(self) -> None:
        """Clear document/chat state while retaining sign-in and request-ID deduplication."""
        self.document = None
        self.source = None
        self.proposal = None
        self.modified = False
        self.chat.clear()
        self.research_draft = False
        self.save_error = None


class SessionStore:
    """Thread-safe registry of at most four sessions with a one-hour inactivity expiry."""

    def __init__(self, settings: Settings) -> None:
        """Initialize an empty session registry sharing the application's configuration."""
        self.settings = settings
        self.sessions: dict[str, Session] = {}
        self.lock = threading.Lock()

    def get(self, cookie: str | None) -> Session:
        """Refresh an existing session or create one after evicting inactive entries."""
        with self.lock:
            now = time.monotonic()
            expired = [key for key, session in self.sessions.items() if now - session.touched > 3600]
            for key in expired:
                del self.sessions[key]
            if cookie not in self.sessions:
                if len(self.sessions) >= 4:
                    raise DemoError("The four-session demo limit was reached. Reuse an existing browser session.", 503)
                key = uuid4().hex
                self.sessions[key] = Session(id=key, drive=OneDrive(self.settings))
                cookie = key
            assert cookie is not None
            session = self.sessions[cookie]
            session.touched = now
            return session


class VersionRequest(BaseModel):
    """Request guarded by the SHA-256 version of the current working document."""

    version: str


class GenerateRequest(VersionRequest):
    """Direct edit instruction and optional selected block IDs for a known snapshot."""

    instruction: str = Field(min_length=1, max_length=2000)
    selected_block_ids: list[str] = Field(default_factory=list, max_length=50)


class ApplyRequest(VersionRequest):
    """Reviewed proposal identity and choice of same-file save or a separate new file."""

    proposal_id: str
    save_as_new: bool = False


class OpenRequest(BaseModel):
    """Identifier of an ordinary DOCX to open from the signed-in user's drive."""

    item_id: str = Field(min_length=1, max_length=512)


class ChatRequest(BaseModel):
    """Chat message with an idempotency key, snapshot/selection, and research consent."""

    message: str = Field(min_length=1, max_length=2000)
    request_id: str = Field(min_length=1, max_length=64, pattern=r"^[a-zA-Z0-9-]+$")
    version: str | None = None
    selected_block_ids: list[str] = Field(default_factory=list, max_length=50)
    allow_research: bool = False
    folder_id: str | None = Field(default=None, max_length=512)


class CreateRequest(VersionRequest):
    """Unsaved research-draft version and optional destination folder for a save retry."""

    folder_id: str | None = Field(default=None, max_length=512)


def _session(request: Request) -> Session:
    """Resolve the session assigned to an API request by the local-protection middleware."""
    return request.state.demo_session


SessionDependency = Annotated[Session, Depends(_session)]


def _document(session: Session, version: str | None = None) -> WordDocument:
    """Require a loaded document and, when provided, its exact current version."""
    if session.document is None:
        raise DemoError("Load a Word document first.", 409)
    if version is not None and session.document.version != version:
        raise DemoError("The document changed. Refresh the preview and generate a new proposal.", 409)
    return session.document


def _view(session: Session) -> dict[str, object]:
    """Serialize the working document, preview, source identity, and save-recovery state."""
    document = _document(session)
    return {
        "name": document.name, "version": document.version,
        "blocks": [block.model_dump() for block in document.blocks],
        "nodes": [node.model_dump() for node in document.nodes()],
        "warnings": document.warnings, "modified": session.modified,
        "source": session.source.model_dump() if session.source else None,
        "research_draft": session.research_draft,
        "save_error": session.save_error,
    }


def _proposal_view(proposal: Proposal) -> dict[str, object]:
    """Serialize review changes and the staged candidate without applying any edits."""
    return {
        "id": proposal.id, "version": proposal.version, "summary": proposal.plan.summary,
        "changes": [change.model_dump() for change in proposal.changes],
        "preview": {
            "nodes": [node.model_dump() for node in proposal.candidate.nodes()],
            "warnings": proposal.candidate.warnings,
        } if proposal.candidate else None,
    }


def _chat_view(session: Session, receipt: dict[str, object]) -> dict[str, object]:
    """Combine lightweight request metadata with the session's current, not cached, view."""
    return {
        **receipt,
        "chat": [turn.model_dump() for turn in session.chat],
        "document": _view(session) if session.document else None,
        "proposal": _proposal_view(session.proposal) if session.proposal and session.proposal.changes else None,
    }


def _load(session: Session, data: bytes, name: str, source: SourceReference | None = None) -> dict[str, object]:
    """Validate a new DOCX before replacing working state and clearing its old transcript."""
    document = WordDocument(data, name)
    session.clear_document()
    session.document = document
    session.source = source
    return _view(session)


def _update_source(session: Session) -> GraphItem:
    """Save current bytes to the opened item and advance source metadata after success."""
    if session.document is None or session.source is None:
        raise DemoError("Open a OneDrive document before saving.", 409)
    item = session.drive.update(session.document.data, session.source)
    session.source = session.source.model_copy(update={
        "etag": item.eTag, "name": item.name,
        "web_url": item.webUrl or session.source.web_url,
    })
    return item


def create_app(settings: Settings | None = None) -> FastAPI:
    """Create the local-only app, session store, cloud/agent clients, routes, and static mount."""
    settings = settings or Settings.from_env()
    application = FastAPI(title="Word editing showcase", docs_url=None, redoc_url=None)
    application.add_middleware(TrustedHostMiddleware, allowed_hosts=["127.0.0.1", "localhost"])
    store = SessionStore(settings)
    application.state.store = store
    application.state.editor = AzureEditor(settings)
    application.state.team = AgentTeam(settings)

    @application.exception_handler(DemoError)
    async def demo_error(request: Request, exc: DemoError) -> JSONResponse:
        """Return a safe domain error and log its route/status without document content."""
        logger.warning("Demo request failed: %s (%s)", request.url.path, exc.status)
        return JSONResponse({"error": exc.message}, status_code=exc.status)

    @application.exception_handler(RequestValidationError)
    async def validation_error(request: Request, exc: RequestValidationError) -> JSONResponse:
        """Return a generic validation error without echoing invalid request data."""
        logger.warning("Invalid request: %s", request.url.path)
        return JSONResponse({"error": "Invalid request fields or input size. Refresh and retry."}, status_code=422)

    @application.exception_handler(Exception)
    async def unexpected_error(request: Request, exc: Exception) -> JSONResponse:
        """Log an unexpected exception type and return a non-sensitive HTTP 500 response."""
        logger.error("Unexpected %s in %s", type(exc).__name__, request.url.path)
        return JSONResponse({"error": "An unexpected server error occurred. Check the server console for the error type."}, status_code=500)

    @application.middleware("http")
    async def protect_local_api(request: Request, call_next):
        """Guard API mutations, assign sessions, and set cache and browser-security headers."""
        if request.url.path.startswith("/api/"):
            if request.method not in ("GET", "HEAD"):
                origin = request.headers.get("Origin")
                expected = f"{request.url.scheme}://{request.headers.get('host', '')}"
                if (
                    request.headers.get("X-Demo-Request") != "1"
                    or (origin is not None and origin != expected)
                    or request.headers.get("Sec-Fetch-Site") == "cross-site"
                ):
                    return JSONResponse({"error": "Only same-origin demo requests are allowed."}, status_code=403)
                length = request.headers.get("Content-Length")
                if length and (not length.isdigit() or int(length) > MAX_FILE_BYTES + 65536):
                    return JSONResponse({"error": "The request exceeds the 10 MiB demo limit."}, status_code=413)
            try:
                session = store.get(request.cookies.get(COOKIE))
            except DemoError as exc:
                return JSONResponse({"error": exc.message}, status_code=exc.status)
            request.state.demo_session = session
            response = await call_next(request)
            response.set_cookie(COOKIE, session.id, httponly=True, samesite="strict", max_age=3600)
            response.headers["Cache-Control"] = "no-store"
        else:
            response = await call_next(request)
            if request.url.path == "/" or request.url.path.startswith("/static/"):
                response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "no-referrer"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self'; style-src 'self'; "
            "img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; "
            "base-uri 'none'; form-action 'self'"
        )
        return response

    @application.get("/")
    def index() -> FileResponse:
        """Serve the browser entry page from the application's packaged static files."""
        return FileResponse(Path(__file__).parent / "static" / "index.html")

    @application.get("/api/status")
    def status(session: SessionDependency) -> dict[str, object]:
        """Return setup status and recoverable document, chat, and proposal state."""
        with session.lock:
            return {
                "graph_problem": settings.graph_problem(),
                "ai_problem": settings.ai_problem(),
                "ai_auth": settings.ai_auth,
                "connected": session.drive.connected(),
                "document": _view(session) if session.document else None,
                "chat": [turn.model_dump() for turn in session.chat],
                "proposal": _proposal_view(session.proposal) if session.proposal else None,
            }

    @application.post("/api/auth/login")
    def login(session: SessionDependency) -> dict[str, bool]:
        """Complete interactive delegated sign-in while serializing session mutations."""
        with session.lock:
            session.drive.sign_in()
            return {"connected": True}

    @application.get("/api/auth/pending")
    def pending_login(session: SessionDependency) -> dict[str, str | None]:
        """Expose the current sign-in link without waiting on the sign-in-held session lock."""
        return {"url": session.drive.login_url}

    @application.post("/api/auth/logout")
    def logout(session: SessionDependency) -> dict[str, bool]:
        """Remove cached sign-in accounts and clear the session's document and transcript."""
        with session.lock:
            session.drive.sign_out()
            session.clear_document()
            return {"connected": False}

    @application.get("/api/drive")
    def browse(session: SessionDependency, folder_id: str | None = None, cursor: str | None = None) -> dict[str, object]:
        """Validate query sizes and return a supported page of the user's OneDrive entries."""
        if (folder_id and len(folder_id) > 512) or (cursor and len(cursor) > 8192):
            raise DemoError("The OneDrive folder request is too long.")
        with session.lock:
            entries, next_link = session.drive.browse(folder_id, cursor)
            return {"entries": [entry.model_dump() for entry in entries], "next": next_link}

    @application.post("/api/document/onedrive")
    def open_onedrive(body: OpenRequest, session: SessionDependency) -> dict[str, object]:
        """Download a cloud DOCX and load it with the identity/eTag needed for later saves."""
        with session.lock:
            data, source = session.drive.download(body.item_id)
            return _load(session, data, source.name, source)

    @application.post("/api/document/sample")
    def sample(session: SessionDependency) -> dict[str, object]:
        """Load the bundled synthetic DOCX without needing a cloud connection."""
        path = ROOT / "demo" / "sample.docx"
        if not path.is_file():
            raise DemoError("The sample fixture is missing. Run npm install and npm run sample.", 503)
        with session.lock:
            return _load(session, path.read_bytes(), "Northstar-proposal.docx")

    @application.post("/api/document/upload")
    def upload(session: SessionDependency, file: UploadFile = File(...)) -> dict[str, object]:
        """Read a size-bounded local upload and validate it before replacing working state."""
        data = file.file.read(MAX_FILE_BYTES + 1)
        with session.lock:
            return _load(session, data, (file.filename or "upload.docx").replace("\\", "/").rsplit("/", 1)[-1])

    @application.post("/api/document/reset")
    def reset(session: SessionDependency) -> dict[str, bool]:
        """Clear the working document and chat while preserving the drive connection."""
        with session.lock:
            session.clear_document()
            return {"cleared": True}

    @application.post("/api/proposal")
    def propose(body: GenerateRequest, session: SessionDependency) -> dict[str, object]:
        """Stage a direct Azure proposal after snapshot, selection, and document validation."""
        if not body.instruction.strip():
            raise DemoError("Enter an editing instruction.")
        with session.lock:
            document = _document(session, body.version)
            selected = set(body.selected_block_ids)
            if selected - {block.id for block in document.blocks if block.editable}:
                raise DemoError("The selection contains unknown or read-only blocks. Select again.")
            instruction = body.instruction
            if selected:
                instruction += "\n\nEDIT SCOPE: Change only these selected block IDs: " + ", ".join(body.selected_block_ids)
            plan = application.state.editor.propose(document, instruction)
            targets = (
                {edit.block_id for edit in plan.edits} | {edit.block_id for edit in plan.formats}
                | {block_id for table in plan.tables for block_id in table.block_ids}
            )
            if selected and targets - selected:
                raise DemoError("The proposal changes blocks outside your selection. Nothing was applied; narrow the instruction and retry.")
            changes = document.preview_plan(plan)
            candidate = document.apply_plan(plan) if changes else None
            proposal = Proposal(
                id=uuid4().hex, version=document.version, plan=plan, changes=changes, candidate=candidate,
            )
            session.proposal = proposal if changes else None
            return _proposal_view(proposal)

    @application.post("/api/proposal/apply")
    def apply(body: ApplyRequest, session: SessionDependency) -> dict[str, object]:
        """Install the reviewed candidate and save to the same cloud item or a new copy.

        Local documents remain local. Expected cloud-save failures retain applied bytes
        and report save_error, leaving download/retry recovery available. Stale proposal
        IDs or versions are rejected before applying or saving.
        """
        with session.lock:
            document = _document(session, body.version)
            proposal = session.proposal
            if proposal is None or proposal.id != body.proposal_id or proposal.version != document.version:
                raise DemoError("This proposal is no longer current. Generate a new one.", 409)
            if proposal.candidate is None:
                raise DemoError("This proposal contains no changes.", 409)
            session.document = proposal.candidate
            session.proposal = None
            session.modified = True
            session.save_error = None
            message = "Applied the reviewed changes locally. Download the updated document."
            if session.source is not None:
                try:
                    if body.save_as_new:
                        item = session.drive.save_copy(session.document.data, session.source)
                        session.source = SourceReference(
                            item_id=item.id, parent_id=session.source.parent_id, name=item.name,
                            etag=item.eTag, web_url=item.webUrl,
                        )
                        message = f"Applied and saved as a new OneDrive file: {item.name}. The original was not updated."
                    else:
                        _update_source(session)
                        message = f"Applied and saved to the same OneDrive file: {session.source.name}."
                    session.modified = False
                except DemoError as exc:
                    logger.warning("Apply cloud save failed (%s)", exc.status)
                    session.save_error = exc.message
                    message = "Changes applied locally, but the OneDrive save was not confirmed. " + exc.message
            elif body.save_as_new:
                session.save_error = "This is a local document. Download it; open a OneDrive document to save a cloud copy."
                message += " " + session.save_error
            if session.chat:
                session.chat.append(ChatTurn(role="assistant", text=message, agent="Word tools"))
            return {**_view(session), "save_error": session.save_error, "apply_message": message}

    @application.post("/api/proposal/discard")
    def discard(session: SessionDependency) -> dict[str, bool]:
        """Drop the unapplied proposal without changing the working document."""
        with session.lock:
            session.proposal = None
            if session.chat:
                session.chat.append(ChatTurn(role="assistant", text="Discarded the proposal. The working Word document is unchanged.", agent="Word tools"))
            return {"discarded": True}

    @application.get("/api/chat/activity")
    def chat_activity(session: SessionDependency) -> dict[str, str]:
        """Read live workflow progress without blocking on the chat-held document lock."""
        # A chat holds the document lock; activity polling must not wait on it.
        return {"activity": session.activity}

    @application.post("/api/chat")
    def chat(body: ChatRequest, session: SessionDependency) -> dict[str, object]:
        """Run one conversational workflow and retain its reply, draft, or staged proposal.

        Deduplicate successful request IDs and return current session state on replay.
        Research creates a new OneDrive report; uncertain saves retain a local draft.
        Edit results remain unapplied until review, and workflow activity is always reset.
        """
        if not body.message.strip():
            raise DemoError("Enter a chat message.")
        with session.lock:
            if body.request_id in session.chat_results:
                return _chat_view(session, session.chat_results[body.request_id])
            if session.document is not None:
                _document(session, body.version)
                if body.version is None:
                    raise DemoError("Refresh the current document before chatting.", 409)
            elif body.version is not None:
                raise DemoError("The document is no longer loaded. Refresh state.", 409)
            editable_ids = {b.id for b in session.document.blocks if b.editable} if session.document else set()
            if set(body.selected_block_ids) - editable_ids:
                raise DemoError("The selection is stale or read-only. Select the blocks again.", 409)
            connected = session.drive.connected()
            if body.allow_research and not connected:
                raise DemoError("Connect Microsoft 365 before asking for research. Research reports are saved automatically to OneDrive.", 409)
            stages: list[str] = []

            def progress(text: str) -> None:
                """Publish the latest workflow stage and retain it in this request's receipt."""
                session.activity = text
                stages.append(text)

            try:
                result = asyncio.run(application.state.team.run(
                    body.message, session.chat, session.document, body.selected_block_ids,
                    body.allow_research, progress, research_connected=connected,
                    pending_plan=session.proposal.plan if session.proposal else None,
                ))
                save_error: str | None = None
                if result.action == "research":
                    if not connected:
                        raise DemoError("Connect Microsoft 365 before asking for research. Research reports are saved automatically to OneDrive.", 409)
                    if session.modified:
                        raise DemoError("Save/download or clear your existing unsaved document before replacing it with a research report.", 409)
                    if result.document is None:
                        raise DemoError("The research workflow did not create a document.", 502)
                    session.document, session.source = result.document, None
                    session.modified, session.research_draft = True, True
                    session.save_error = None
                    session.proposal = None
                    progress("Saving your research report to OneDrive")
                    try:
                        session.source = session.drive.create_document(
                            result.document.data, result.document.name, body.folder_id,
                        )
                        session.modified, session.research_draft = False, False
                        result.message = f"Saved {session.source.name} to OneDrive. What would you like to change? Edits are reviewed before saving back."
                    except DemoError as exc:
                        save_error = exc.message
                        session.save_error = exc.message
                        logger.warning("Research report cloud creation failed (%s)", exc.status)
                        result.message = "Research draft created locally, but the OneDrive save was not confirmed. Download the draft and check OneDrive before retrying."
                elif result.action == "edit":
                    document = _document(session, body.version)
                    if result.plan is None:
                        raise DemoError("The editor did not return a proposal.", 502)
                    changes = document.preview_plan(result.plan)
                    # The reviewed candidate comes from the isolated MCP validator.
                    if changes and result.document is None:
                        raise DemoError("The editor did not return a validated candidate. Nothing was changed.", 502)
                    session.proposal = Proposal(
                        id=uuid4().hex, version=document.version, plan=result.plan,
                        changes=changes, candidate=result.document,
                    ) if changes else None
                    if changes:
                        result.message += " Review and apply it, or tell me what to change in the proposal."
                session.chat.extend([
                    ChatTurn(role="user", text=body.message),
                    ChatTurn(role="assistant", text=result.message, agent=(
                        "Researcher + Report writer + Word MCP" if result.action == "research"
                        else "Word editor + Word MCP" if result.action == "edit" else "Coordinator"
                    )),
                ])
                receipt = {
                    "stages": stages, "sources": [source.model_dump() for source in result.sources],
                    "save_error": save_error, "research_saved": result.action == "research" and save_error is None,
                }
                session.chat_results[body.request_id] = receipt
                return _chat_view(session, receipt)
            finally:
                session.activity = "Ready"

    @application.post("/api/chat/clear")
    def clear_chat(session: SessionDependency) -> dict[str, bool]:
        """Clear transcript entries without clearing the document, proposal, or receipts."""
        with session.lock:
            session.chat.clear()
            return {"cleared": True}

    @application.post("/api/document/create")
    def create_cloud_document(body: CreateRequest, session: SessionDependency) -> dict[str, object]:
        """Retry creating a new unsaved research draft without rerunning web research."""
        with session.lock:
            document = _document(session, body.version)
            if session.source is not None or not session.research_draft:
                raise DemoError("Only a new unsaved research draft can be created with this action.", 409)
            session.source = session.drive.create_document(document.data, document.name, body.folder_id)
            session.modified, session.research_draft = False, False
            session.save_error = None
            return _view(session)

    @application.get("/api/document/download")
    def download(session: SessionDependency) -> Response:
        """Return current DOCX bytes with a safe attachment name and no-store caching."""
        with session.lock:
            document = _document(session)
            name = document.name[:-5] + ("-edited" if session.modified else "") + ".docx"
            return Response(document.data, media_type=DOCX_TYPE, headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{quote(name, safe='')}",
                "Cache-Control": "no-store",
            })

    @application.post("/api/document/save")
    def save(body: VersionRequest, session: SessionDependency) -> dict[str, str | None]:
        """Create an additional cloud sibling without changing the active source reference."""
        with session.lock:
            document = _document(session, body.version)
            if session.source is None:
                raise DemoError("Load a OneDrive document to save a sibling copy. Local documents can be downloaded.", 409)
            item = session.drive.save_copy(document.data, session.source)
            return {"name": item.name, "web_url": item.webUrl}

    @application.post("/api/document/update")
    def update(body: VersionRequest, session: SessionDependency) -> dict[str, object]:
        """Retry saving applied unsaved bytes to the same item with its cloud-version guard."""
        with session.lock:
            document = _document(session, body.version)
            if session.source is None:
                raise DemoError("Open a OneDrive document first. Local files can only be downloaded.", 409)
            if not session.modified:
                raise DemoError("There are no applied edits to save.", 409)
            item = _update_source(session)
            session.modified = False
            session.save_error = None
            return {"name": item.name, "web_url": item.webUrl, "document": _view(session)}

    application.mount("/static", StaticFiles(directory=Path(__file__).parent / "static"), name="static")
    return application


app = create_app()
