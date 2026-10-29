from datetime import datetime, timezone
import re
import socket
from typing import Literal
from urllib.parse import quote, urlsplit
from uuid import uuid4

import httpx
import msal
from msal.oauth2cli.oauth2 import BrowserInteractionTimeoutError
from pydantic import BaseModel, Field, ValidationError
from requests.exceptions import RequestException

from app.config import Settings
from app.documents import MAX_FILE_BYTES
from app.errors import DemoError

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["Files.ReadWrite"]


class ParentReference(BaseModel):
    id: str


class GraphItem(BaseModel):
    id: str
    name: str
    size: int = 0
    webUrl: str | None = None
    file: dict[str, object] | None = None
    folder: dict[str, object] | None = None
    remoteItem: dict[str, object] | None = None
    parentReference: ParentReference | None = None
    eTag: str | None = None


class GraphPage(BaseModel):
    value: list[GraphItem]
    next_link: str | None = Field(default=None, alias="@odata.nextLink")


class DriveEntry(BaseModel):
    id: str
    name: str
    kind: Literal["folder", "docx"]
    size: int


class SourceReference(BaseModel):
    item_id: str
    parent_id: str
    name: str
    web_url: str | None = None
    etag: str | None = None


def _cloud_url(url: str, *, graph: bool = False) -> str:
    try:
        parsed = urlsplit(url)
        host = parsed.hostname or ""
        valid = (
            parsed.scheme == "https" and not parsed.username and not parsed.password
            and parsed.port in (None, 443) and not parsed.fragment
        )
    except ValueError as exc:
        raise DemoError("Microsoft returned an invalid file URL.", 502) from exc
    if graph:
        valid = valid and host == "graph.microsoft.com" and parsed.path.startswith("/v1.0/")
    else:
        valid = valid and (
            host == "graph.microsoft.com" or any(host.endswith(suffix) for suffix in (
                ".sharepoint.com", ".sharepoint-df.com", ".1drv.com",
                ".onedrive.com", ".blob.core.windows.net", ".storage.live.com",
            ))
        )
    if not valid:
        raise DemoError("Microsoft returned an unsupported file URL; no request was sent.", 502)
    return url


def _check(response: httpx.Response) -> None:
    if response.is_success:
        return
    status = response.status_code
    if status == 401:
        raise DemoError("Your Microsoft sign-in expired. Sign in again.", 401)
    if status == 403:
        raise DemoError("OneDrive access was denied. Check delegated Files.ReadWrite consent and folder permissions.", 403)
    if status == 404:
        raise DemoError("The OneDrive item no longer exists or is not accessible.", 404)
    if status == 409:
        raise DemoError("A copy with that name already exists. Retry to use a new name; nothing was overwritten.", 409)
    if status == 412:
        raise DemoError("The OneDrive document changed since you opened it. Download your working copy, then reopen the cloud file before saving.", 409)
    if status == 429:
        retry = response.headers.get("Retry-After", "")
        seconds = f" Wait {retry} seconds." if retry.isdigit() else " Wait before retrying."
        raise DemoError("OneDrive is throttling requests." + seconds, 429)
    raise DemoError(f"OneDrive returned HTTP {status}. The operation was not confirmed.", 502)


class OneDrive:
    def __init__(self, settings: Settings, transport: httpx.BaseTransport | None = None) -> None:
        self.settings = settings
        self.transport = transport
        self._client: msal.PublicClientApplication | None = None
        self.login_url: str | None = None

    def sign_in(self) -> None:
        problem = self.settings.graph_problem()
        if problem:
            raise DemoError(problem, 503)
        try:
            with socket.socket() as listener:
                listener.bind(("127.0.0.1", 0))
                port = listener.getsockname()[1]
            client = msal.PublicClientApplication(
                self.settings.client_id,
                authority=f"https://login.microsoftonline.com/{self.settings.tenant_id}",
            )
            def launching(**kwargs: object) -> None:
                self.login_url = f"http://localhost:{port}?welcome=true"

            result = client.acquire_token_interactive(
                scopes=SCOPES, timeout=180, prompt="select_account", port=port,
                on_before_launching_ui=launching,
                welcome_template=(
                    "<!doctype html><html><head><title>Microsoft 365 sign-in</title></head>"
                    '<body><h1>Connect Wordsmith to OneDrive</h1>'
                    '<p><a href="$auth_uri">Continue to Microsoft sign-in</a></p>'
                    "<p>Use the account containing your Test folder. After signing in, return to Wordsmith.</p>"
                    "</body></html>"
                ),
                auth_uri_callback=lambda uri: setattr(self, "login_url", uri),
            )
        except BrowserInteractionTimeoutError as exc:
            raise DemoError(
                "Microsoft sign-in timed out after three minutes. Click Connect again, then use "
                "the visible Open Microsoft sign-in link and finish signing in on this computer.",
                408,
            ) from exc
        except (RequestException, ValueError) as exc:
            raise DemoError("Microsoft sign-in could not start. Check the tenant/client IDs, loopback redirect, and network.", 502) from exc
        except OSError as exc:
            raise DemoError("Could not start the local sign-in listener. Close the previous sign-in window and retry.", 503) from exc
        finally:
            self.login_url = None
        if "access_token" not in result:
            code = str(result.get("error", "sign_in_failed"))
            safe_code = re.sub(r"[^a-zA-Z0-9_]", "", code)[:60]
            raise DemoError(
                f"Microsoft sign-in did not complete ({safe_code}). Check consent, tenant policy, and the http://localhost desktop redirect.",
                401,
            )
        self._client = client

    def sign_out(self) -> None:
        if self._client is not None:
            for account in self._client.get_accounts():
                self._client.remove_account(account)
        self._client = None

    def connected(self) -> bool:
        return self._client is not None and bool(self._client.get_accounts())

    def _token(self) -> str:
        if self._client is None:
            raise DemoError("Sign in to Microsoft 365 first.", 401)
        accounts = self._client.get_accounts()
        try:
            result = self._client.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        except RequestException as exc:
            raise DemoError("Microsoft token renewal failed. Check your network and sign in again.", 504) from exc
        if not result or "access_token" not in result:
            raise DemoError("Your Microsoft sign-in needs attention. Sign in again.", 401)
        return str(result["access_token"])

    def _http(self) -> httpx.Client:
        return httpx.Client(transport=self.transport, timeout=45, follow_redirects=False)

    def _get_json(self, path: str) -> object:
        url = _cloud_url(path if path.startswith("https://") else GRAPH + path, graph=True)
        try:
            with self._http() as client:
                response = client.get(url, headers={"Authorization": "Bearer " + self._token()})
                _check(response)
                return response.json()
        except httpx.HTTPError as exc:
            raise DemoError("Could not reach OneDrive. Check your network and retry.", 504) from exc
        except ValueError as exc:
            raise DemoError("OneDrive returned an unreadable response.", 502) from exc

    def browse(self, folder_id: str | None, cursor: str | None) -> tuple[list[DriveEntry], str | None]:
        path = cursor or (
            f"/me/drive/items/{quote(folder_id, safe='')}/children?$top=100"
            if folder_id else "/me/drive/root/children?$top=100"
        )
        try:
            page = GraphPage.model_validate(self._get_json(path))
        except ValidationError as exc:
            raise DemoError("OneDrive returned an unexpected folder response.", 502) from exc
        entries = [
            DriveEntry(
                id=item.id, name=item.name,
                kind="folder" if item.folder is not None else "docx", size=item.size,
            )
            for item in page.value
            if item.remoteItem is None and (
                item.folder is not None or (item.file is not None and item.name.lower().endswith(".docx"))
            )
        ]
        entries.sort(key=lambda item: (item.kind != "folder", item.name.casefold()))
        return entries, page.next_link

    def download(self, item_id: str) -> tuple[bytes, SourceReference]:
        escaped = quote(item_id, safe="")
        try:
            item = GraphItem.model_validate(self._get_json(f"/me/drive/items/{escaped}"))
        except ValidationError as exc:
            raise DemoError("OneDrive returned unexpected document metadata.", 502) from exc
        if item.file is None or item.remoteItem is not None or not item.name.lower().endswith(".docx"):
            raise DemoError("Choose an ordinary .docx in your own OneDrive.")
        if item.size > MAX_FILE_BYTES:
            raise DemoError("This demo accepts Word files up to 10 MiB.", 413)
        if item.parentReference is None:
            raise DemoError("OneDrive did not identify this document's folder.", 502)
        url = GRAPH + f"/me/drive/items/{escaped}/content"
        try:
            with self._http() as client:
                for _ in range(4):
                    _cloud_url(url)
                    headers = {"Authorization": "Bearer " + self._token()} if urlsplit(url).hostname == "graph.microsoft.com" else {}
                    with client.stream("GET", url, headers=headers) as response:
                        if response.status_code in (301, 302, 303, 307, 308):
                            location = response.headers.get("Location")
                            if not location:
                                raise DemoError("OneDrive returned a redirect without a download URL.", 502)
                            url = location
                            continue
                        _check(response)
                        chunks: list[bytes] = []
                        size = 0
                        for chunk in response.iter_bytes(65536):
                            size += len(chunk)
                            if size > MAX_FILE_BYTES:
                                raise DemoError("The downloaded file exceeds the 10 MiB limit.", 413)
                            chunks.append(chunk)
                        if item.eTag:
                            current = GraphItem.model_validate(self._get_json(f"/me/drive/items/{escaped}"))
                            if current.eTag != item.eTag:
                                raise DemoError("The OneDrive document changed while downloading. Reopen it before editing.", 409)
                        return b"".join(chunks), SourceReference(
                            item_id=item.id, parent_id=item.parentReference.id,
                            name=item.name, web_url=item.webUrl, etag=item.eTag,
                        )
                raise DemoError("OneDrive returned too many download redirects.", 502)
        except httpx.HTTPError as exc:
            raise DemoError("The OneDrive download failed. Check your network and retry.", 504) from exc
        except ValidationError as exc:
            raise DemoError("OneDrive returned invalid document-version metadata. Reopen the file.", 502) from exc

    def update(self, data: bytes, source: SourceReference) -> GraphItem:
        if not source.etag:
            raise DemoError("No cloud version is available for a safe update. Reopen this document from OneDrive.", 409)
        path = f"/me/drive/items/{quote(source.item_id, safe='')}/content"
        try:
            with self._http() as client:
                response = client.put(
                    GRAPH + path, content=data,
                    headers={
                        "Authorization": "Bearer " + self._token(),
                        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "If-Match": source.etag,
                    },
                )
                _check(response)
                if response.status_code not in (200, 201):
                    raise DemoError("The update was not confirmed. Check OneDrive before retrying.", 502)
                item = GraphItem.model_validate(response.json())
                if item.id != source.item_id:
                    raise DemoError("OneDrive did not confirm the expected document update. Check OneDrive before retrying.", 502)
                return item
        except httpx.HTTPError as exc:
            raise DemoError("The update could not be confirmed. Check OneDrive before retrying; your working copy is still available.", 504) from exc
        except (ValueError, ValidationError) as exc:
            raise DemoError("OneDrive returned an unexpected update response. Check OneDrive before retrying.", 502) from exc

    def save_copy(self, data: bytes, source: SourceReference) -> GraphItem:
        stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
        stem = re.sub(r'[\\/:*?"<>|]', "-", source.name[:-5])[:100]
        name = f"{stem}-edited-{stamp}-{uuid4().hex[:8]}.docx"
        item = self._create_upload(data, source.parent_id, name)
        if item.id == source.item_id:
            raise DemoError("OneDrive did not confirm a distinct copy. Check OneDrive before retrying.", 502)
        return item

    def create_document(self, data: bytes, name: str, folder_id: str | None) -> SourceReference:
        folder_path = f"/me/drive/items/{quote(folder_id, safe='')}" if folder_id else "/me/drive/root"
        try:
            folder = GraphItem.model_validate(self._get_json(folder_path))
        except ValidationError as exc:
            raise DemoError("OneDrive returned invalid destination-folder metadata.", 502) from exc
        if folder.folder is None or folder.remoteItem is not None:
            raise DemoError("Choose an ordinary folder in your own OneDrive.", 422)
        stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "-", name.removesuffix(".docx")).strip(" .")[:100] or "Research report"
        unique_name = f"{stem}-{uuid4().hex[:8]}.docx"
        item = self._create_upload(data, folder.id, unique_name)
        return SourceReference(
            item_id=item.id, parent_id=folder.id, name=item.name, web_url=item.webUrl, etag=item.eTag,
        )

    def _create_upload(self, data: bytes, parent_id: str, name: str) -> GraphItem:
        if not data or len(data) > MAX_FILE_BYTES:
            raise DemoError("Cloud document creation requires a nonempty file under 10 MiB.", 413)
        path = f"/me/drive/items/{quote(parent_id, safe='')}:/{quote(name, safe='')}:/createUploadSession"
        try:
            with self._http() as client:
                response = client.post(
                    GRAPH + path,
                    headers={"Authorization": "Bearer " + self._token()},
                    json={"item": {"@microsoft.graph.conflictBehavior": "fail", "name": name}},
                )
                _check(response)
                session = response.json()
                if not isinstance(session, dict) or not isinstance(session.get("uploadUrl"), str):
                    raise DemoError("OneDrive did not return an upload session.", 502)
                upload_url = _cloud_url(session["uploadUrl"])
                response = client.put(
                    upload_url, content=data,
                    headers={
                        "Content-Type": "application/octet-stream",
                        "Content-Length": str(len(data)),
                        "Content-Range": f"bytes 0-{len(data) - 1}/{len(data)}",
                    },
                )
                _check(response)
                if response.status_code not in (200, 201):
                    raise DemoError("The upload was not confirmed complete. Check OneDrive before retrying.", 502)
                item = GraphItem.model_validate(response.json())
                return item
        except httpx.HTTPError as exc:
            raise DemoError("The save could not be confirmed. Check OneDrive before retrying; the original was not targeted.", 504) from exc
        except (ValueError, ValidationError) as exc:
            raise DemoError("OneDrive returned an unexpected upload response. Check OneDrive before retrying.", 502) from exc
