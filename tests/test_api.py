from fastapi.testclient import TestClient
import pytest

from app.config import Settings
from app.documents import Edit, EditPlan, FormatEdit, ParagraphFormat, TableEdit, WordDocument
from app.main import COOKIE, create_app
from app.onedrive import GraphItem, SourceReference
from app.errors import DemoError


class TestEditor:
    __test__ = False

    def propose(self, document, instruction):
        block = next(b for b in document.blocks if "30 days" in b.text)
        return EditPlan(summary="Update payment terms.", edits=[
            Edit(block_id=block.id, original="30", replacement="45"),
        ])


@pytest.fixture
def client():
    app = create_app(Settings())
    with TestClient(app, base_url="http://localhost", headers={"X-Demo-Request": "1"}) as client:
        yield client


def session(client):
    return client.app.state.store.sessions[client.cookies.get(COOKIE)]


def test_missing_configuration_and_sample_download(client, sample_bytes):
    status = client.get("/api/status").json()
    assert status["graph_problem"] and status["ai_problem"]
    assert not status["connected"]
    result = client.post("/api/document/sample", json={})
    assert result.status_code == 200
    assert result.json()["blocks"]
    response = client.get("/api/document/download")
    assert response.content == sample_bytes
    assert ".docx" in response.headers["content-disposition"]
    assert response.headers["cache-control"] == "no-store"
    assert client.cookies.get(COOKIE)


def test_browser_assets_revalidate_and_use_matching_revision(client):
    import re
    response = client.get("/")
    assert response.headers["cache-control"] == "no-cache"
    assets = re.findall(r'(?:src|href)="(/static/[^"]+)"', response.text)
    assert len(assets) == 3
    revisions = {asset.split("?v=")[1] for asset in assets}
    assert len(revisions) == 1
    for asset in assets:
        assert client.get(asset).headers["cache-control"] == "no-cache"


def test_real_model_configuration_is_required(client):
    view = client.post("/api/document/sample", json={}).json()
    response = client.post("/api/proposal", json={"version": view["version"], "instruction": "Update terms."})
    assert response.status_code == 503
    assert "AZURE_OPENAI" in response.json()["error"]


def test_proposal_review_discard_apply_and_stale_version(client, sample_bytes):
    client.app.state.editor = TestEditor()
    view = client.post("/api/document/sample", json={}).json()
    request = {"version": view["version"], "instruction": "Update payment terms."}
    proposal = client.post("/api/proposal", json=request).json()
    assert proposal["changes"][0]["before"] != proposal["changes"][0]["after"]
    assert client.get("/api/document/download").content == sample_bytes
    assert client.post("/api/proposal/discard", json={}).status_code == 200
    assert client.get("/api/document/download").content == sample_bytes
    apply = {"version": view["version"], "proposal_id": proposal["id"]}
    assert client.post("/api/proposal/apply", json=apply).status_code == 409
    proposal = client.post("/api/proposal", json=request).json()
    apply["proposal_id"] = proposal["id"]
    updated = client.post("/api/proposal/apply", json=apply)
    assert updated.status_code == 200
    assert updated.json()["modified"]
    assert any("45 days" in block["text"] for block in updated.json()["blocks"])
    assert client.post("/api/proposal", json=request).status_code == 409
    assert client.post("/api/proposal/apply", json=apply).status_code == 409


def test_legacy_empty_proposal_is_not_staged(client):
    class NoopEditor:
        def propose(self, document, instruction):
            return EditPlan(summary="Unsupported operation", edits=[])
    client.app.state.editor = NoopEditor()
    view = client.post("/api/document/sample", json={}).json()
    response = client.post("/api/proposal", json={"version": view["version"], "instruction": "Unsupported"})
    assert response.status_code == 200 and response.json()["changes"] == []
    assert session(client).proposal is None
    assert client.get("/api/status").json()["proposal"] is None


def test_formatting_and_table_proposal_preview_is_exact_applied_docx(client, sample_bytes):
    writes = []
    class RichDrive:
        def connected(self): return True
        def download(self, item_id):
            return sample_bytes, SourceReference(
                item_id="original", parent_id="test", name="rich.docx", etag='"v1"',
            )
        def update(self, data, source):
            writes.append((data, source.item_id, source.etag))
            return GraphItem(id="original", name="rich.docx", eTag='"v2"')
    class RichEditor:
        def propose(self, document, instruction):
            blocks = [b for b in document.blocks if b.kind == "list"]
            return EditPlan(summary="Blue headings and scope table.", edits=[], formats=[
                FormatEdit(block_id=b.id, original=b.text, format=ParagraphFormat(color="0000FF"))
                for b in document.blocks if b.kind == "heading"
            ], tables=[TableEdit(
                block_ids=[b.id for b in blocks], originals=[b.text for b in blocks],
                rows=[["Scope"], *[[b.text] for b in blocks]], header=True,
            )])
    client.get("/api/status")
    session(client).drive = RichDrive()
    client.app.state.editor = RichEditor()
    view = client.post("/api/document/onedrive", json={"item_id": "original"}).json()
    response = client.post("/api/proposal", json={"version": view["version"], "instruction": "Format it"})
    assert response.status_code == 200
    proposal = response.json()
    assert {change["kind"] for change in proposal["changes"]} == {"format", "table"}
    assert sum(node["kind"] == "table" for node in proposal["preview"]["nodes"]) == 2
    assert client.get("/api/document/download").content == sample_bytes
    assert writes == []
    candidate_bytes = session(client).proposal.candidate.data
    applied = client.post("/api/proposal/apply", json={"version": view["version"], "proposal_id": proposal["id"]}).json()
    assert applied["nodes"] == proposal["preview"]["nodes"]
    assert not applied["modified"]
    assert client.get("/api/document/download").content == candidate_bytes
    assert len(WordDocument(candidate_bytes, "result.docx").body.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl")) == 2
    assert writes == [(candidate_bytes, "original", '"v1"')]
    assert session(client).source.etag == '"v2"'


def test_selection_limits_model_targets_and_rejects_unknown_blocks(client, sample_bytes):
    client.app.state.editor = TestEditor()
    view = client.post("/api/document/sample", json={}).json()
    request = {"version": view["version"], "instruction": "Change terms.", "selected_block_ids": ["unknown"]}
    assert client.post("/api/proposal", json=request).status_code == 400
    request["selected_block_ids"] = [view["blocks"][0]["id"]]
    response = client.post("/api/proposal", json=request)
    assert response.status_code == 400
    assert "outside your selection" in response.json()["error"]
    assert client.get("/api/document/download").content == sample_bytes
    request["selected_block_ids"] = [next(b["id"] for b in view["blocks"] if "30 days" in b["text"])]
    assert client.post("/api/proposal", json=request).status_code == 200


def test_empty_plan_explains_unsupported_requests_and_cannot_apply(client, sample_bytes):
    class EmptyEditor:
        def propose(self, document, instruction):
            return EditPlan(summary="This operation is unsupported.", edits=[])
    client.app.state.editor = EmptyEditor()
    view = client.post("/api/document/sample", json={}).json()
    proposal = client.post("/api/proposal", json={"version": view["version"], "instruction": "Add a chart."}).json()
    assert proposal["preview"] is None and proposal["changes"] == []
    assert "unsupported" in proposal["summary"]
    assert client.post("/api/proposal/apply", json={"version": view["version"], "proposal_id": proposal["id"]}).status_code == 409
    assert client.get("/api/document/download").content == sample_bytes


def test_upload_invalid_request_and_clear(client, sample_bytes):
    assert client.post("/api/document/upload", files={"file": ("sample.docx", sample_bytes)}).status_code == 200
    assert client.post("/api/document/upload", files={"file": ("not.docx", b"invalid")}).status_code == 400
    assert client.get("/api/status").json()["document"] is not None
    assert client.post("/api/document/reset", json={}).status_code == 200
    assert client.get("/api/document/download").status_code == 409


def test_local_source_cannot_be_saved_to_onedrive(client):
    view = client.post("/api/document/sample", json={}).json()
    assert client.post("/api/document/save", json={"version": view["version"]}).status_code == 409


def test_save_explicitly_creates_copy_only(client, sample_bytes):
    calls = []

    class FakeDrive:
        def connected(self): return True
        def download(self, item_id):
            return sample_bytes, SourceReference(item_id="original", parent_id="folder", name="source.docx")
        def update(self, data, source):
            raise DemoError("Missing cloud version; save a copy instead.", 409)
        def save_copy(self, data, source):
            calls.append((data, source.item_id))
            return GraphItem(id="copy", name="source-edited.docx", webUrl="https://example.sharepoint.com/copy")

    client.get("/api/status")
    session(client).drive = FakeDrive()
    client.app.state.editor = TestEditor()
    view = client.post("/api/document/onedrive", json={"item_id": "original"}).json()
    proposal = client.post("/api/proposal", json={"version": view["version"], "instruction": "Change terms."}).json()
    assert calls == []
    view = client.post("/api/proposal/apply", json={"version": view["version"], "proposal_id": proposal["id"]}).json()
    assert calls == []
    result = client.post("/api/document/save", json={"version": view["version"]})
    assert result.status_code == 200
    assert result.json()["name"] == "source-edited.docx"
    assert len(calls) == 1 and calls[0][1] == "original"
    assert calls[0][0] != sample_bytes


def test_origin_and_host_guards(client):
    assert client.post("/api/document/sample", json={}, headers={"Origin": "https://evil.example"}).status_code == 403
    assert client.post("/api/document/sample", json={}, headers={"X-Demo-Request": ""}).status_code == 403
    assert client.get("/api/status", headers={"Host": "evil.example"}).status_code == 400
    assert client.post("/api/document/sample", content=b"", headers={"Content-Length": str(11 * 1024 * 1024)}).status_code == 413


def test_ai_requires_nonblank_instruction_and_session_isolation(client):
    view = client.post("/api/document/sample", json={}).json()
    assert client.post("/api/proposal", json={"version": view["version"], "instruction": "   "}).status_code == 400
    with TestClient(client.app, base_url="http://localhost") as other:
        assert other.get("/api/status").json()["document"] is None


def test_pending_signin_can_be_read_while_document_lock_is_held(client):
    from concurrent.futures import ThreadPoolExecutor
    client.get("/api/status")
    current = session(client)
    current.drive.login_url = "http://localhost:12345?welcome=true"
    with current.lock, ThreadPoolExecutor() as pool:
        response = pool.submit(client.get, "/api/auth/pending").result(timeout=3)
    assert response.status_code == 200
    assert response.json()["url"] == current.drive.login_url
    assert response.headers["cache-control"] == "no-store"


def test_apply_saves_same_document_and_advances_cloud_version(client, sample_bytes):
    calls = []
    class FakeDrive:
        def connected(self): return True
        def download(self, item_id):
            return sample_bytes, SourceReference(item_id="original", parent_id="test", name="test.docx", etag='"1"')
        def update(self, data, source):
            calls.append(source.etag)
            return GraphItem(id="original", name="test.docx", eTag='"2"', webUrl="https://example.sharepoint.com/test.docx")
    client.get("/api/status")
    session(client).drive = FakeDrive()
    client.app.state.editor = TestEditor()
    view = client.post("/api/document/onedrive", json={"item_id": "original"}).json()
    assert client.post("/api/document/update", json={"version": view["version"]}).status_code == 409
    proposal = client.post("/api/proposal", json={"version": view["version"], "instruction": "Change terms."}).json()
    view = client.post("/api/proposal/apply", json={"version": view["version"], "proposal_id": proposal["id"]}).json()
    assert calls == ['"1"']
    assert not view["modified"] and view["save_error"] is None
    result = client.post("/api/document/update", json={"version": view["version"]})
    assert result.status_code == 409
    assert calls == ['"1"']
    assert session(client).source.etag == '"2"'


def test_failed_update_preserves_working_copy(client):
    class FailedDrive:
        def update(self, data, source):
            raise DemoError("The OneDrive document changed since you opened it.", 409)
    view = client.post("/api/document/sample", json={}).json()
    current = session(client)
    current.drive = FailedDrive()
    current.source = SourceReference(item_id="original", parent_id="test", name="test.docx", etag='"1"')
    current.modified = True
    before = current.document.data
    result = client.post("/api/document/update", json={"version": view["version"]})
    assert result.status_code == 409
    assert current.modified and current.document.data == before and current.source.etag == '"1"'


@pytest.mark.parametrize("copy", [False, True])
@pytest.mark.parametrize("fail", [False, True])
def test_apply_cloud_destination_failure_recovery_and_no_replay(client, copy, fail):
    writes = []
    class Drive:
        def connected(self): return True
        def update(self, data, source):
            writes.append(("same", data, source.item_id, source.etag))
            if fail:
                raise DemoError("Cloud version conflict", 409)
            return GraphItem(id=source.item_id, name=source.name, eTag='"2"')
        def save_copy(self, data, source):
            writes.append(("copy", data, source.item_id, source.etag))
            if fail:
                raise DemoError("Copy save not confirmed", 504)
            return GraphItem(id="new", name="new.docx", eTag='"new-1"', webUrl="https://example.com/new")
    view = client.post("/api/document/sample", json={}).json()
    current = session(client)
    current.drive = Drive()
    current.source = SourceReference(item_id="original", parent_id="test", name="test.docx", etag='"1"')
    client.app.state.editor = TestEditor()
    proposal = client.post("/api/proposal", json={"version": view["version"], "instruction": "Change terms"}).json()
    candidate = current.proposal.candidate.data
    request = {"version": view["version"], "proposal_id": proposal["id"], "save_as_new": copy}
    response = client.post("/api/proposal/apply", json=request)
    assert response.status_code == 200
    result = response.json()
    assert writes == [("copy" if copy else "same", candidate, "original", '"1"')]
    assert client.get("/api/document/download").content == candidate
    assert current.proposal is None
    assert result["modified"] is fail
    assert bool(result["save_error"]) is fail
    assert client.get("/api/status").json()["document"]["save_error"] == result["save_error"]
    assert current.source.item_id == ("new" if copy and not fail else "original")
    assert current.source.etag == ('"1"' if fail else '"new-1"' if copy else '"2"')
    assert client.post("/api/proposal/apply", json=request).status_code == 409
    assert len(writes) == 1
