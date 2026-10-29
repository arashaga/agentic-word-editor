from types import SimpleNamespace
import json

import httpx
import pytest

from app import ai
from app.ai import AzureEditor
from app.config import Settings
from app.documents import Edit, EditPlan, FormatEdit, ParagraphFormat
from app.errors import DemoError
from app.onedrive import GRAPH, OneDrive, SourceReference
from app import onedrive
from msal.oauth2cli.oauth2 import BrowserInteractionTimeoutError


def test_sign_in_timeout_is_actionable_and_clears_link(monkeypatch):
    drive = OneDrive(Settings(
        tenant_id="11111111-1111-1111-1111-111111111111",
        client_id="22222222-2222-2222-2222-222222222222",
    ))
    class Client:
        def __init__(self, *args, **kwargs): pass
        def acquire_token_interactive(self, **kwargs):
            kwargs["on_before_launching_ui"](ui="browser")
            assert drive.login_url.startswith("http://localhost:")
            assert "$auth_uri" in kwargs["welcome_template"]
            raise BrowserInteractionTimeoutError("synthetic timeout")
    monkeypatch.setattr(onedrive.msal, "PublicClientApplication", Client)
    with pytest.raises(DemoError, match="visible Open Microsoft sign-in link") as raised:
        drive.sign_in()
    assert raised.value.status == 408
    assert drive.login_url is None
    assert not drive.connected()


def test_successful_sign_in_keeps_client_and_clears_link(monkeypatch):
    drive = OneDrive(Settings(
        tenant_id="11111111-1111-1111-1111-111111111111",
        client_id="22222222-2222-2222-2222-222222222222",
    ))
    class Client:
        def __init__(self, *args, **kwargs): pass
        def acquire_token_interactive(self, **kwargs):
            kwargs["on_before_launching_ui"](ui="browser")
            assert drive.login_url
            return {"access_token": "synthetic"}
        def get_accounts(self): return [{"home_account_id": "synthetic"}]
    monkeypatch.setattr(onedrive.msal, "PublicClientApplication", Client)
    drive.sign_in()
    assert drive.connected()
    assert drive.login_url is None


@pytest.mark.parametrize("status", [200, 412, 423])
def test_update_sends_cloud_version_and_handles_conflicts(monkeypatch, sample_bytes, status):
    requests = []
    def handler(request):
        requests.append(request)
        assert request.method == "PUT"
        assert request.url.path == "/v1.0/me/drive/items/original/content"
        assert request.headers["If-Match"] == '"version-1"'
        assert request.content == sample_bytes
        return httpx.Response(status, json={"id": "original", "name": "demo.docx", "eTag": '"version-2"'})
    drive = drive_with(handler, monkeypatch)
    source = SourceReference(item_id="original", parent_id="test", name="demo.docx", etag='"version-1"')
    if status == 200:
        assert drive.update(sample_bytes, source).eTag == '"version-2"'
    else:
        with pytest.raises(DemoError, match="changed since" if status == 412 else "not confirmed"):
            drive.update(sample_bytes, source)
    assert len(requests) == 1


def test_update_without_cloud_version_never_sends_request(monkeypatch, sample_bytes):
    def handler(request):
        pytest.fail("No network request should be sent without a cloud version")
    drive = drive_with(handler, monkeypatch)
    with pytest.raises(DemoError, match="safe update"):
        drive.update(sample_bytes, SourceReference(item_id="original", parent_id="test", name="demo.docx"))


@pytest.mark.parametrize("folder_id", ["test-folder", None])
def test_create_report_uses_new_name_conflict_fail_and_returns_editable_source(monkeypatch, sample_bytes, folder_id):
    seen = []
    def handler(request):
        seen.append(request)
        if request.method == "GET":
            assert request.url.path.endswith("/items/test-folder" if folder_id else "/root")
            return httpx.Response(200, json={"id": "test-folder", "name": "Test", "folder": {}})
        if request.method == "POST":
            payload = json.loads(request.content)
            assert payload["item"]["@microsoft.graph.conflictBehavior"] == "fail"
            assert payload["item"]["name"].startswith("Inflation-")
            assert payload["item"]["name"].endswith(".docx")
            assert "/items/test-folder:" in request.url.path
            return httpx.Response(200, json={"uploadUrl": "https://example.sharepoint.com/upload"})
        assert request.method == "PUT" and request.url.host == "example.sharepoint.com"
        assert "Authorization" not in request.headers
        assert request.content == sample_bytes
        return httpx.Response(201, json={"id": "created", "name": "Inflation-unique.docx", "eTag": '"v1"', "file": {}})
    source = drive_with(handler, monkeypatch).create_document(sample_bytes, "Inflation.docx", folder_id)
    assert source.item_id == "created" and source.parent_id == "test-folder" and source.etag == '"v1"'
    assert [r.method for r in seen] == ["GET", "POST", "PUT"]


def test_create_report_rejects_nonfolder_without_upload(monkeypatch, sample_bytes):
    seen = []
    def handler(request):
        seen.append(request)
        return httpx.Response(200, json={"id": "not-folder", "name": "Document.docx", "file": {}})
    with pytest.raises(DemoError, match="ordinary folder"):
        drive_with(handler, monkeypatch).create_document(sample_bytes, "Inflation.docx", "not-folder")
    assert len(seen) == 1


def drive_with(handler, monkeypatch):
    drive = OneDrive(Settings(), transport=httpx.MockTransport(handler))
    monkeypatch.setattr(drive, "_token", lambda: "test-only-token")
    return drive


def test_graph_pagination_and_external_cursor_guard(monkeypatch):
    next_link = GRAPH + "/me/drive/root/children?$skiptoken=next"
    requests = []

    def handler(request):
        requests.append(request)
        assert request.headers["Authorization"] == "Bearer test-only-token"
        return httpx.Response(200, json={
            "value": [
                {"id": "folder", "name": "Folder", "folder": {}},
                {"id": "doc", "name": "Proposal.docx", "file": {}, "size": 20},
                {"id": "pdf", "name": "Other.pdf", "file": {}},
            ], "@odata.nextLink": next_link,
        })

    drive = drive_with(handler, monkeypatch)
    entries, cursor = drive.browse(None, None)
    assert [entry.kind for entry in entries] == ["folder", "docx"]
    assert cursor == next_link
    drive.browse(None, cursor)
    assert len(requests) == 2
    with pytest.raises(DemoError, match="unsupported"):
        drive.browse(None, "https://evil.example/steal-token")
    assert len(requests) == 2


def test_graph_redirect_does_not_forward_bearer(monkeypatch, sample_bytes):
    def handler(request):
        if request.url.path.endswith("/content"):
            assert request.headers["Authorization"]
            return httpx.Response(302, headers={"Location": "https://example.sharepoint.com/download?temporary=synthetic"})
        if request.url.host == "example.sharepoint.com":
            assert "Authorization" not in request.headers
            return httpx.Response(200, content=sample_bytes)
        return httpx.Response(200, json={
            "id": "source", "name": "sample.docx", "file": {}, "size": len(sample_bytes),
            "parentReference": {"id": "folder"},
        })

    data, source = drive_with(handler, monkeypatch).download("source")
    assert data == sample_bytes and source.parent_id == "folder"


def test_download_rejects_concurrent_cloud_change(monkeypatch, sample_bytes):
    metadata_reads = 0
    def handler(request):
        nonlocal metadata_reads
        if request.url.path.endswith("/content"):
            return httpx.Response(200, content=sample_bytes)
        metadata_reads += 1
        return httpx.Response(200, json={
            "id": "source", "name": "sample.docx", "file": {}, "size": len(sample_bytes),
            "parentReference": {"id": "test"}, "eTag": f'"version-{metadata_reads}"',
        })
    with pytest.raises(DemoError, match="changed while downloading"):
        drive_with(handler, monkeypatch).download("source")


def test_graph_copy_uses_new_name_conflict_fail_and_no_bearer_to_upload(monkeypatch, sample_bytes):
    paths = []

    def handler(request):
        paths.append(request.url.path)
        if request.method == "POST":
            assert "/items/folder:/" in request.url.path
            assert "/items/original/" not in request.url.path
            assert "/items/original/createUploadSession" not in request.url.path
            assert b'"@microsoft.graph.conflictBehavior":"fail"' in request.content
            assert request.headers["Authorization"]
            return httpx.Response(200, json={"uploadUrl": "https://example.sharepoint.com/upload"})
        assert request.method == "PUT"
        assert "Authorization" not in request.headers
        assert request.content == sample_bytes
        assert request.headers["Content-Range"] == f"bytes 0-{len(sample_bytes)-1}/{len(sample_bytes)}"
        return httpx.Response(201, json={"id": "copy", "name": "edited.docx"})

    drive = drive_with(handler, monkeypatch)
    result = drive.save_copy(sample_bytes, SourceReference(item_id="original", parent_id="folder", name="original.docx"))
    assert result.id == "copy" and len(paths) == 2


@pytest.mark.parametrize("status,text", [(401, "expired"), (403, "denied"), (404, "no longer"), (409, "nothing was overwritten"), (429, "throttling"), (503, "not confirmed")])
def test_graph_failures_are_explicit(monkeypatch, status, text):
    drive = drive_with(lambda request: httpx.Response(status, json={}), monkeypatch)
    with pytest.raises(DemoError, match=text):
        drive.browse(None, None)


def test_unconfirmed_upload_is_not_success(monkeypatch, sample_bytes):
    def handler(request):
        return httpx.Response(200, json={"uploadUrl": "https://example.sharepoint.com/upload"}) if request.method == "POST" else httpx.Response(202, json={})
    drive = drive_with(handler, monkeypatch)
    with pytest.raises(DemoError, match="not confirmed complete"):
        drive.save_copy(sample_bytes, SourceReference(item_id="original", parent_id="folder", name="original.docx"))


def test_copy_conflict_does_not_upload_or_overwrite(monkeypatch, sample_bytes):
    calls = []
    def handler(request):
        calls.append(request.method)
        return httpx.Response(409, json={})
    drive = drive_with(handler, monkeypatch)
    with pytest.raises(DemoError, match="nothing was overwritten"):
        drive.save_copy(sample_bytes, SourceReference(item_id="original", parent_id="folder", name="original.docx"))
    assert calls == ["POST"]


def test_graph_timeout_is_explicit(monkeypatch):
    def handler(request):
        raise httpx.ReadTimeout("synthetic timeout", request=request)
    with pytest.raises(DemoError, match="network"):
        drive_with(handler, monkeypatch).browse(None, None)


def mock_model(monkeypatch, message=None, error=None):
    captured = {}

    class Client:
        def __init__(self, **kwargs):
            captured.update(kwargs)
            self.beta = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(parse=self.parse)))
        def __enter__(self): return self
        def __exit__(self, *args): return False
        def parse(self, **kwargs):
            captured["request"] = kwargs
            if error: raise error
            return SimpleNamespace(choices=[SimpleNamespace(message=message)])

    monkeypatch.setattr(ai, "OpenAI", Client)
    return captured


def editor():
    return AzureEditor(Settings(
        ai_endpoint="https://example.openai.azure.com", ai_deployment="configured-deployment",
        ai_auth="key", ai_key="synthetic-key-for-tests",
    ))


def test_azure_proposal_validates_and_uses_configured_endpoint(monkeypatch, document):
    block = next(b for b in document.blocks if "30 days" in b.text)
    plan = EditPlan(summary="Payment update.", edits=[Edit(block_id=block.id, original="30", replacement="45")])
    captured = mock_model(monkeypatch, SimpleNamespace(parsed=plan, refusal=None))
    assert editor().propose(document, "Change payment terms.") == plan
    assert captured["base_url"] == "https://example.openai.azure.com/openai/v1/"
    assert captured["request"]["model"] == "configured-deployment"
    assert captured["request"]["response_format"] is EditPlan
    assert captured["max_retries"] == 0


def test_azure_receives_formatting_context_and_accepts_formatting_only(monkeypatch, document):
    block = next(b for b in document.blocks if b.kind == "heading")
    plan = EditPlan(summary="Blue heading", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(color="0000FF")),
    ])
    captured = mock_model(monkeypatch, SimpleNamespace(parsed=plan, refusal=None))
    assert editor().propose(document, "Make the heading blue") == plan
    content = json.loads(captured["request"]["messages"][1]["content"])
    context = next(b for b in content["document_blocks"] if b["block_id"] == block.id)
    assert context["kind"] == "heading" and context["can_convert"]
    assert context["format"]["font_size_pt"] == 16
    assert "color" in context["format"]


def test_structured_schema_requires_nullable_format_fields_and_all_operation_arrays():
    from openai.lib._pydantic import to_strict_json_schema
    schema = to_strict_json_schema(EditPlan)
    assert set(schema["required"]) == {"summary", "edits", "formats", "tables"}
    for model in schema["$defs"].values():
        if model.get("type") == "object":
            assert model["additionalProperties"] is False
            assert set(model["required"]) == set(model["properties"])


@pytest.mark.parametrize("repair", [True, False])
def test_legacy_editor_rechecks_empty_plan_once(monkeypatch, document, repair):
    requests = []
    empty = EditPlan(summary="Color is unsupported.", edits=[])
    block = next(b for b in document.blocks if b.kind == "heading")
    color = EditPlan(summary="Blue heading", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(color="0000FF")),
    ])
    class Client:
        def __init__(self, **kwargs):
            self.beta = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(parse=self.parse)))
        def __enter__(self): return self
        def __exit__(self, *args): return False
        def parse(self, **kwargs):
            requests.append(list(kwargs["messages"]))
            value = color if repair and len(requests) == 2 else empty
            return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(parsed=value, refusal=None))])
    monkeypatch.setattr(ai, "OpenAI", Client)
    result = editor().propose(document, "Make the heading blue.")
    assert len(requests) == 2
    assert requests[1][-1]["content"] == ai.EMPTY_PLAN_RECHECK
    assert result == (color if repair else empty)


def test_azure_refusal_and_invalid_block_are_not_success(monkeypatch, document):
    mock_model(monkeypatch, SimpleNamespace(parsed=None, refusal="synthetic refusal"))
    with pytest.raises(DemoError, match="declined"):
        editor().propose(document, "An instruction")
    mock_model(monkeypatch, SimpleNamespace(
        parsed=EditPlan(summary="Invalid", edits=[Edit(block_id="made-up", original="30", replacement="45")]),
        refusal=None,
    ))
    with pytest.raises(DemoError, match="unknown"):
        editor().propose(document, "An instruction")


@pytest.mark.parametrize("error_type,expected", [(ai.RateLimitError, "rate-limiting"), (ai.AuthenticationError, "authentication failed"), (ai.BadRequestError, "structured outputs")])
def test_azure_api_errors_are_actionable(monkeypatch, document, error_type, expected):
    response = httpx.Response(400, request=httpx.Request("POST", "https://example.openai.azure.com"))
    mock_model(monkeypatch, error=error_type("synthetic error", response=response, body=None))
    with pytest.raises(DemoError, match=expected):
        editor().propose(document, "An instruction")


def test_azure_timeout_is_not_success(monkeypatch, document):
    request = httpx.Request("POST", "https://example.openai.azure.com")
    mock_model(monkeypatch, error=ai.APITimeoutError(request=request))
    with pytest.raises(DemoError, match="in time"):
        editor().propose(document, "An instruction")


@pytest.mark.parametrize("endpoint", ["http://localhost:9000", "https://evil.example", "https://example.openai.azure.com:bad"])
def test_invalid_azure_endpoints_are_configuration_errors(endpoint):
    assert Settings(ai_endpoint=endpoint, ai_deployment="model").ai_problem()


@pytest.mark.parametrize("endpoint,expected", [
    ("https://example.services.ai.azure.com/openai/v1/responses", "Remove the trailing responses"),
    ("https://example.services.ai.azure.com/api/projects/demo", "not the Foundry project endpoint"),
])
def test_wrong_endpoint_kind_has_actionable_error(endpoint, expected):
    settings = Settings(ai_endpoint=endpoint, ai_deployment="model")
    assert expected in settings.ai_problem()
    with pytest.raises(DemoError, match=expected):
        settings.ai_base_url()


@pytest.mark.parametrize("endpoint", [
    "https://example.services.ai.azure.com/openai/v1/",
    "https://example.openai.azure.com",
])
def test_passwordless_resource_endpoint_does_not_require_a_key(endpoint):
    settings = Settings(ai_endpoint=endpoint, ai_deployment="model", ai_auth="entra")
    assert settings.ai_problem() is None
    assert settings.ai_base_url().endswith("/openai/v1/")
