"""Tests for research reports, agent routing, and the Word MCP workflow."""

import asyncio
import base64
import json

import httpx
from openai import AsyncOpenAI
from agent_framework import AgentResponse, Content, Message
from agent_framework.exceptions import ChatClientException
from lxml import etree
import pytest

from app import agents
from app.agents import AgentResult, AgentTeam, ChatTurn, Route, research_sources, word_tools
from app.config import Settings
from app.documents import EditPlan, FormatEdit, ParagraphFormat, W
from app.errors import DemoError
from app.reports import ResearchReport, ResearchSource, ReportSection, create_report


@pytest.fixture
def sources():
    """Provide one synthetic public citation for report-generation tests."""
    return [ResearchSource(title="Synthetic economic release", url="https://www.bls.gov/example")]


@pytest.fixture
def report():
    """Build a fictional report whose two sections cite the fixture source."""
    return ResearchReport(title="Synthetic inflation report", subtitle="Illustrative evidence, not live economic data.", sections=[
        ReportSection(heading="Findings", paragraphs=["The fictional index is illustrative."], source_ids=[1]),
        ReportSection(heading="Causes and caveats", paragraphs=["No causal inference is made."], source_ids=[1]),
    ])


def test_report_is_real_word_with_citations_and_editable_headings(report, sources):
    """Verify a report exports as an editable Word package with citations and layout."""
    document = create_report(report, sources)
    assert document.data.startswith(b"PK")
    assert document.name == "Synthetic inflation report.docx"
    assert any(block.text == "Sources: [1]" for block in document.blocks)
    assert any(block.text == sources[0].url for block in document.blocks)
    assert any(block.kind == "heading" and block.text == "Findings" for block in document.blocks)
    assert all(block.editable for block in document.blocks)
    section = document.body.find(W + "sectPr")
    assert section.find(W + "pgSz").get(W + "w") == "12240"
    assert etree.fromstring(document.parts["word/settings.xml"]).find(W + "zoom").get(W + "percent") == "100"


@pytest.mark.parametrize("case", ["missing", "unknown", "bad_url", "bad_text", "too_big"])
def test_invalid_reports_are_rejected(report, sources, case):
    """Reject reports with missing sources, unsafe URLs, invalid text, or excessive size."""
    if case == "missing": sources = []
    elif case == "unknown": report.sections[0].source_ids = [2]
    elif case == "bad_url": sources[0].url = "javascript:alert(1)"
    elif case == "bad_text": report.sections[0].paragraphs = ["bad\x00text"]
    elif case == "too_big": report.sections[0].paragraphs = ["x" * 25000]
    with pytest.raises(DemoError):
        create_report(report, sources)


def test_only_provider_citations_are_accepted(sources):
    """Keep only valid provider annotations and reject prose-invented citations."""
    response = AgentResponse(messages=[Message("assistant", [
        Content("text", text="Brief", annotations=[
            {"type": "citation", "title": sources[0].title, "url": sources[0].url},
            {"type": "citation", "title": sources[0].title, "url": sources[0].url},
            {"type": "citation", "title": "Bad", "url": "javascript:alert(1)"},
        ]),
    ])])
    assert research_sources(response) == sources
    with pytest.raises(DemoError, match="no usable source"):
        research_sources(AgentResponse(messages=[Message("assistant", ["Invented https://example.com"])]))


def test_real_mcp_protocol_edit_create_export_and_isolation(document, report, sources):
    """Exercise document loading, formatting preview, report creation, export, and isolation."""
    async def exercise():
        """Run the MCP operations against one tool session and check a second is empty."""
        async with word_tools() as tools:
            await tools.call("load_document", data_base64=base64.b64encode(document.data).decode(), name=document.name)
            context = await tools.call("inspect_document")
            assert context["version"] == document.version
            block = document.blocks[0]
            plan = EditPlan(summary="Blue heading", edits=[], formats=[
                FormatEdit(block_id=block.id, original=block.text,
                           format=ParagraphFormat(color="0000FF", font_family="Libertinus Serif")),
            ])
            preview = await tools.call("preview_edits", plan_json=plan.model_dump_json(), version=document.version)
            assert preview["changes"][0]["kind"] == "format"
            updated = await tools.export()
            assert updated.version != document.version
            assert all(run.format.color == "0000FF" for run in updated.blocks[0].runs)
            assert all(run.format.font_family == "Libertinus Serif" for run in updated.blocks[0].runs)
            with pytest.raises(DemoError, match="snapshot changed"):
                await tools.call("preview_edits", plan_json=plan.model_dump_json(), version="stale")
            with pytest.raises(DemoError, match="No validated candidate"):
                await tools.export()
            await tools.call("create_research_report", report_json=report.model_dump_json(),
                             sources_json=json.dumps([s.model_dump() for s in sources]))
            created = await tools.export()
            assert any(block.text == "Findings" for block in created.blocks)
        async with word_tools() as separate:
            with pytest.raises(DemoError, match="Load a document"):
                await separate.call("inspect_document")
    asyncio.run(exercise())


def fake_agents(monkeypatch, route, report=None, sources=None, fail_at=None):
    """Install role-aware fake agents and return a log of their inputs."""
    calls = []

    class FakeAgent:
        """Mimic coordinator, researcher, and writer responses for team tests."""
        def __init__(self, *, name, **kwargs):
            """Record the framework-assigned role name."""
            self.name = name
        async def run(self, message, **kwargs):
            """Log a role invocation and return its configured synthetic response."""
            calls.append((self.name, message, kwargs))
            if self.name == fail_at:
                raise ChatClientException("synthetic provider failure")
            if self.name == "Coordinator":
                return AgentResponse(value=route)
            if self.name == "WebResearcher":
                return AgentResponse(messages=[Message("assistant", [
                    Content("text", text="A fictional public brief.", annotations=[
                        {"type": "citation", "title": s.title, "url": s.url} for s in sources
                    ]),
                ])])
            if self.name == "ReportWriter":
                return AgentResponse(value=report)
            raise AssertionError(self.name)
    monkeypatch.setattr(agents, "Agent", FakeAgent)
    return calls


def team():
    """Create an AgentTeam configured with synthetic Azure OpenAI credentials."""
    return AgentTeam(Settings(ai_endpoint="https://example.openai.azure.com", ai_deployment="test", ai_auth="key", ai_key="synthetic"))


def test_team_research_uses_framework_roles_and_real_mcp(monkeypatch, report, sources):
    """Verify research follows coordinator, researcher, and writer stages through MCP."""
    calls = fake_agents(monkeypatch, Route(action="research", task="Public synthetic economics", reply=""), report, sources)
    stages = []
    result = asyncio.run(team().run("Research it", [], None, [], True, stages.append, research_connected=True))
    assert result.document and result.action == "research"
    assert [name for name, *_ in calls] == ["Coordinator", "WebResearcher", "ReportWriter"]
    assert result.sources == sources
    assert "Public synthetic economics" in calls[1][1]
    assert any("Word MCP" in stage for stage in stages)


def test_team_does_not_send_document_to_researcher(monkeypatch, document, report, sources):
    """Ensure research receives neither document contents nor prior editing conversation."""
    calls = fake_agents(monkeypatch, Route(action="research", task="Public synthetic economics", reply=""), report, sources)
    result = asyncio.run(team().run("Research public inflation", [ChatTurn(role="user", text="Prior edit")],
                                   document, [], True, lambda _: None, research_connected=True))
    assert result.document
    for name, message, _ in calls:
        if name == "WebResearcher":
            assert document.name not in message and "30 days" not in message and "Prior edit" not in message


def test_research_requires_consent_and_no_fake_fallback(monkeypatch):
    """Require public-web consent before research and do not synthesize a fallback report."""
    calls = fake_agents(monkeypatch, Route(action="research", task="Inflation", reply=""))
    with pytest.raises(DemoError, match="Enable public web"):
        asyncio.run(team().run("Research", [], None, [], False, lambda _: None, research_connected=True))
    assert [call[0] for call in calls] == ["Coordinator"]


def test_provider_failure_is_actionable_after_mcp_cleanup(monkeypatch):
    """Surface provider failure as an actionable error after cleaning up MCP resources."""
    fake_agents(monkeypatch, Route(action="research", task="Inflation", reply=""), fail_at="WebResearcher")
    with pytest.raises(DemoError, match="No simulated research"):
        asyncio.run(team().run("Research", [], None, [], True, lambda _: None, research_connected=True))


def test_coordinator_can_clarify_and_uses_conversation_history(monkeypatch):
    """Return coordinator clarification and include prior turns in its routing input."""
    calls = fake_agents(monkeypatch, Route(action="clarify", task="", reply="Which country?"))
    result = asyncio.run(team().run("Research inflation", [ChatTurn(role="user", text="Hello")],
                                   None, [], False, lambda _: None))
    assert result.message == "Which country?" and result.document is None
    assert json.loads(calls[0][1])["history"][0]["text"] == "Hello"


def test_disconnected_research_never_invokes_researcher(monkeypatch):
    """Block research when Microsoft 365 is disconnected before invoking its researcher."""
    calls = fake_agents(monkeypatch, Route(action="research", task="Inflation", reply=""))
    with pytest.raises(DemoError, match="Connect Microsoft 365"):
        asyncio.run(team().run("Research", [], None, [], True, lambda _: None))
    assert [call[0] for call in calls] == ["Coordinator"]


def test_long_history_is_windowed_without_rejecting_conversation(monkeypatch):
    """Bound serialized history by size while retaining the newest conversation turn."""
    calls = fake_agents(monkeypatch, Route(action="clarify", task="", reply="Which country?"))
    history = [ChatTurn(role="user", text=str(index) + "x" * 1000) for index in range(120)]
    asyncio.run(team().run("Continue", history, None, [], False, lambda _: None))
    sent = json.loads(calls[0][1])["history"]
    assert sum(len(turn["text"]) for turn in sent) <= agents.MAX_HISTORY_CHARACTERS
    assert sent[-1]["text"] == history[-1].text
    assert len(sent) < len(history)


@pytest.mark.parametrize("repair", [True, False])
def test_real_framework_schema_and_bounded_empty_plan_recheck(monkeypatch, document, repair):
    """Check structured formatting schema and allow at most one recheck of an empty plan."""
    requests = []
    heading = next(block for block in document.blocks if block.kind == "heading")
    empty = EditPlan(summary="The response schema has no color property.", edits=[])
    color = EditPlan(summary="Blue heading", edits=[], formats=[
        FormatEdit(block_id=heading.id, original=heading.text, format=ParagraphFormat(color="0000FF")),
    ])
    def respond(request):
        """Return a structured route or edit response and capture each model request."""
        body = json.loads(request.content)
        requests.append(body)
        schema = body["response_format"]["json_schema"]
        if schema["name"] == "Route":
            value = Route(action="edit", task="Make the first heading blue.", reply="")
        else:
            properties = schema["schema"]["$defs"]["ParagraphFormat"]["properties"]
            assert {"color", "font_family", "font_size_pt", "bold", "italic", "alignment"} <= properties.keys()
            assert schema["schema"]["$defs"]["FormatEdit"]["properties"]["format"]["$ref"].endswith("/ParagraphFormat")
            value = color if repair and len(requests) == 3 else empty
        return httpx.Response(200, json={
            "id": "test", "object": "chat.completion", "created": 0, "model": "test",
            "choices": [{"index": 0, "finish_reason": "stop",
                         "message": {"role": "assistant", "content": value.model_dump_json()}}],
        })
    async def exercise():
        """Run the real framework client against a local deterministic HTTP transport."""
        async with httpx.AsyncClient(transport=httpx.MockTransport(respond)) as http:
            monkeypatch.setattr(agents, "AsyncOpenAI", lambda **kwargs: AsyncOpenAI(
                api_key="synthetic", base_url="https://example.invalid/v1", http_client=http,
            ))
            return await team().run("Make the first heading blue.", [], document, [], False, lambda _: None)
    result = asyncio.run(exercise())
    assert len(requests) == 3  # Coordinator, initial edit, one recheck. No indefinite retry.
    retry_messages = requests[-1]["messages"]
    assert any("Recheck the empty proposal" in str(message["content"]) for message in retry_messages)
    assert any("formats[i].format.color" in str(message["content"]) for message in retry_messages)
    if repair:
        assert result.document is not None
        assert all(run.format.color == "0000FF" for run in result.document.blocks[0].runs)
    else:
        assert result.document is None and result.plan is not None and not result.plan.formats
