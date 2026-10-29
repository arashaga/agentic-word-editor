"""Private, per-turn stdio MCP worker. No network listener, files, or cloud credentials."""

import base64
import binascii
import json

from mcp.server import MCPServer

from app.documents import EditPlan, MAX_FILE_BYTES, WordDocument
from app.errors import DemoError
from app.reports import ResearchReport, ResearchSource, create_report


def create_server() -> MCPServer:
    server = MCPServer("Wordsmith Word tools", log_level="ERROR")
    document: WordDocument | None = None
    candidate: WordDocument | None = None

    @server.tool()
    def load_document(data_base64: str, name: str) -> dict[str, object]:
        """Host-only: load the current document snapshot into this isolated worker."""
        nonlocal document, candidate
        try:
            if len(data_base64) > (MAX_FILE_BYTES + 2) // 3 * 4:
                raise DemoError("Document payload exceeds the 10 MiB limit.", 413)
            document = WordDocument(base64.b64decode(data_base64, validate=True), name)
            candidate = None
            return {"version": document.version}
        except (DemoError, binascii.Error) as exc:
            return {"error": exc.message if isinstance(exc, DemoError) else "Invalid document encoding."}

    @server.tool()
    def inspect_document() -> dict[str, object]:
        """Read editable blocks, heading kinds, and supported formatting from the loaded DOCX."""
        if document is None:
            return {"error": "Load a document first."}
        try:
            return {"version": document.version, "blocks": [b.model_dump() for b in document.ai_blocks()]}
        except DemoError as exc:
            return {"error": exc.message}

    @server.tool()
    def preview_edits(plan_json: str, version: str) -> dict[str, object]:
        """Validate a structured EditPlan and stage a candidate. Never apply or save it."""
        nonlocal candidate
        candidate = None
        if document is None or document.version != version:
            return {"error": "The document snapshot changed. Generate a new proposal."}
        try:
            plan = EditPlan.model_validate_json(plan_json)
            changes = document.preview_plan(plan)
            candidate = document.apply_plan(plan) if changes else None
            return {"changes": [change.model_dump() for change in changes], "has_candidate": candidate is not None}
        except (DemoError, ValueError) as exc:
            return {"error": exc.message if isinstance(exc, DemoError) else "Invalid edit plan schema."}

    @server.tool()
    def create_research_report(report_json: str, sources_json: str) -> dict[str, object]:
        """Create a real DOCX from a cited structured report; stage locally, never save to cloud."""
        nonlocal candidate
        candidate = None
        try:
            report = ResearchReport.model_validate_json(report_json)
            sources = [ResearchSource.model_validate(source) for source in json.loads(sources_json)]
            candidate = create_report(report, sources)
            return {"name": candidate.name, "version": candidate.version}
        except (DemoError, ValueError, TypeError) as exc:
            return {"error": exc.message if isinstance(exc, DemoError) else "Invalid research report schema."}

    @server.tool()
    def export_candidate() -> dict[str, object]:
        """Host-only: export the validated candidate for review. Does not mutate the source."""
        if candidate is None:
            return {"error": "No validated candidate is available."}
        return {"name": candidate.name, "data_base64": base64.b64encode(candidate.data).decode("ascii")}

    return server


if __name__ == "__main__":
    create_server().run(transport="stdio")
