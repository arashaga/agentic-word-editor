from fastapi.testclient import TestClient
import pytest

from app.agents import AgentResult
from app.config import Settings
from app.documents import EditPlan, FormatEdit, ParagraphFormat, WordDocument
from app.errors import DemoError
from app.main import COOKIE, create_app
from app.onedrive import GraphItem, SourceReference


@pytest.fixture
def client():
    with TestClient(create_app(Settings()), base_url="http://localhost", headers={"X-Demo-Request": "1"}) as client:
        yield client


def session(client):
    return client.app.state.store.sessions[client.cookies.get(COOKIE)]


class Team:
    def __init__(self, sample_bytes):
        self.sample_bytes, self.calls = sample_bytes, []
    async def run(self, message, history, document, selected, consent, progress, **kwargs):
        self.calls.append([message, list(history), consent, kwargs.get("pending_plan")])
        progress("Researcher: synthetic test")
        if message == "clarify":
            return AgentResult(action="clarify", message="Which country?")
        if message == "unsupported":
            return AgentResult(action="edit", message="No edits applied: unsupported operation.",
                               plan=EditPlan(summary="Unsupported operation.", edits=[]))
        if message == "edit":
            block = document.blocks[0]
            plan = EditPlan(summary="Blue heading", edits=[], formats=[
                FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(color="0000FF")),
            ])
            return AgentResult(action="edit", message=plan.summary, document=document.apply_plan(plan), plan=plan)
        return AgentResult(action="research", message="Synthetic report", document=WordDocument(self.sample_bytes, "Research.docx"))


class Drive:
    def __init__(self, fail=False):
        self.writes, self.fail = [], fail
    def connected(self): return True
    def create_document(self, data, name, folder):
        self.writes.append((data, name, folder))
        if self.fail: raise DemoError("Upload not confirmed", 504)
        return SourceReference(item_id="created", parent_id=folder or "root", name=name, etag='"v1"')
    def update(self, data, source):
        self.writes.append((data, source.item_id, source.etag))
        return GraphItem(id=source.item_id, name=source.name, eTag='"v2"')


def test_research_create_followup_review_apply_and_idempotency(client, sample_bytes):
    client.get("/api/status")
    team, drive = Team(sample_bytes), Drive()
    client.app.state.team, session(client).drive = team, drive
    request = {"message": "research and save", "request_id": "first", "allow_research": True,
               "folder_id": "Test"}
    response = client.post("/api/chat", json=request)
    assert response.status_code == 200
    result = response.json()
    assert result["document"]["source"]["item_id"] == "created"
    assert result["document"]["modified"] is False
    assert len(result["chat"]) == 2 and len(drive.writes) == 1
    assert client.post("/api/chat", json=request).json() == result
    assert len(drive.writes) == len(team.calls) == 1
    edit = client.post("/api/chat", json={"message": "edit", "request_id": "second", "version": result["document"]["version"]}).json()
    assert len(team.calls[1][1]) == 2
    assert edit["proposal"]["changes"][0]["kind"] == "format"
    assert client.get("/api/document/download").content == sample_bytes
    assert len(drive.writes) == 1
    client.post("/api/chat/clear", json={})
    client.post("/api/chat", json=request)
    assert len(drive.writes) == 1
    restored = client.get("/api/status").json()
    assert restored["proposal"]["id"] == edit["proposal"]["id"]
    updated = client.post("/api/proposal/apply", json={"version": result["document"]["version"],
                                                    "proposal_id": edit["proposal"]["id"]}).json()
    assert not updated["modified"] and updated["source"]["item_id"] == "created"
    assert updated["source"]["etag"] == '"v2"'
    assert len(drive.writes) == 2


def test_research_draft_can_be_created_later_without_researching_again(client, sample_bytes):
    client.get("/api/status")
    drive, team = Drive(fail=True), Team(sample_bytes)
    session(client).drive, client.app.state.team = drive, team
    result = client.post("/api/chat", json={"message": "research", "request_id": "one"}).json()
    assert result["document"]["research_draft"] and result["document"]["source"] is None
    assert len(drive.writes) == 1
    drive.fail = False
    created = client.post("/api/document/create", json={"version": result["document"]["version"], "folder_id": "Test"})
    assert created.status_code == 200
    assert created.json()["source"]["item_id"] == "created"
    assert not created.json()["research_draft"]
    assert client.post("/api/document/create", json={"version": result["document"]["version"]}).status_code == 409
    assert len(drive.writes) == 2


def test_failed_auto_create_keeps_research_draft_and_reports_error(client, sample_bytes):
    client.get("/api/status")
    drive = Drive(fail=True)
    session(client).drive, client.app.state.team = drive, Team(sample_bytes)
    request = {"message": "research", "request_id": "one"}
    result = client.post("/api/chat", json=request).json()
    assert result["save_error"] == "Upload not confirmed"
    assert result["document"]["research_draft"] and result["document"]["modified"]
    assert result["document"]["source"] is None
    assert client.get("/api/document/download").content == sample_bytes
    client.post("/api/chat", json=request)
    assert len(drive.writes) == 1


def test_stale_inputs_rejected_but_pending_proposal_can_be_revised(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    view = client.post("/api/document/sample", json={}).json()
    request = {"message": "edit", "request_id": "edit", "version": "stale"}
    assert client.post("/api/chat", json=request).status_code == 409
    request["version"] = view["version"]
    request["selected_block_ids"] = ["missing"]
    assert client.post("/api/chat", json=request).status_code == 409
    request["selected_block_ids"] = []
    first = client.post("/api/chat", json=request).json()
    assert first["proposal"]
    request["request_id"] = "next"
    next_result = client.post("/api/chat", json=request).json()
    assert next_result["proposal"]["id"] != first["proposal"]["id"]
    assert client.app.state.team.calls[-1][3].summary == "Blue heading"
    assert client.get("/api/document/download").content == sample_bytes
    assert client.post("/api/proposal/apply", json={"version": view["version"],
                                                  "proposal_id": first["proposal"]["id"]}).status_code == 409


def test_unsaved_document_not_replaced_by_research(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    view = client.post("/api/document/sample", json={}).json()
    session(client).modified = True
    response = client.post("/api/chat", json={"message": "research", "request_id": "one", "version": view["version"]})
    assert response.status_code == 409
    assert client.get("/api/document/download").content == sample_bytes


def test_chat_reset_preserves_document_and_other_sessions_are_isolated(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    view = client.post("/api/document/sample", json={}).json()
    response = client.post("/api/chat", json={"message": "clarify", "request_id": "one", "version": view["version"]})
    assert response.status_code == 200 and len(response.json()["chat"]) == 2
    with TestClient(client.app, base_url="http://localhost") as other:
        assert other.get("/api/status").json()["chat"] == []
    assert client.post("/api/chat/clear", json={}).status_code == 200
    assert client.get("/api/status").json()["chat"] == []
    assert client.get("/api/document/download").content == sample_bytes


def test_noop_does_not_block_followup(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    view = client.post("/api/document/sample", json={}).json()
    request = {"message": "unsupported", "request_id": "no-op", "version": view["version"]}
    result = client.post("/api/chat", json=request).json()
    assert result["proposal"] is None
    assert "apply or discard" not in result["chat"][-1]["text"]
    assert session(client).proposal is None
    request.update(message="edit", request_id="followup")
    assert client.post("/api/chat", json=request).status_code == 200


def test_messages_continue_past_old_session_and_history_caps(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    for index in range(110):
        response = client.post("/api/chat", json={"message": "clarify", "request_id": f"turn-{index}"})
        assert response.status_code == 200, (index, response.text)
    assert len(session(client).chat) == 220
    assert all("document" not in receipt and "chat" not in receipt for receipt in session(client).chat_results.values())


def test_disconnected_research_is_blocked_before_agent_call(client, sample_bytes):
    team = Team(sample_bytes)
    client.app.state.team = team
    response = client.post("/api/chat", json={"message": "research", "request_id": "one", "allow_research": True})
    assert response.status_code == 409 and "Connect Microsoft 365" in response.json()["error"]
    assert team.calls == []


def test_replaying_old_request_returns_current_state_without_rollback(client, sample_bytes):
    client.app.state.team = Team(sample_bytes)
    view = client.post("/api/document/sample", json={}).json()
    request = {"message": "edit", "request_id": "original", "version": view["version"]}
    first = client.post("/api/chat", json=request).json()
    applied = client.post("/api/proposal/apply", json={"version": view["version"],
                                                     "proposal_id": first["proposal"]["id"]}).json()
    replay = client.post("/api/chat", json=request).json()
    assert replay["document"]["version"] == applied["version"]
    assert replay["proposal"] is None
    assert len(client.app.state.team.calls) == 1
