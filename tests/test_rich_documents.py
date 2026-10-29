"""Tests for Word formatting plans, rich previews, and safe list-to-table conversion."""

from io import BytesIO
from zipfile import ZipFile

from lxml import etree
import pytest

from app.documents import Edit, EditPlan, FormatEdit, MAIN_PART, ParagraphFormat, TableEdit, W, WordDocument
from app.errors import DemoError
from conftest import change_xml


def heading_plan(document):
    """Build blue-color formatting operations for every heading block."""
    return EditPlan(summary="Blue headings.", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(color="0000FF"))
        for block in document.blocks if block.kind == "heading"
    ])


def scope_table(document):
    """Describe a header-row table conversion using every existing list block."""
    blocks = [block for block in document.blocks if block.kind == "list"]
    return TableEdit(
        block_ids=[block.id for block in blocks], originals=[block.text for block in blocks],
        rows=[["Scope item"], *[[block.text] for block in blocks]], header=True,
    )


def test_heading_color_changes_real_xml_and_rich_preview_without_changing_text(document):
    """Apply heading colors in OOXML while preserving text and unrelated package parts."""
    plan = heading_plan(document)
    changes = document.preview_plan(plan)
    assert changes and all(change.kind == "format" for change in changes)
    assert all(change.before == change.after for change in changes)
    updated = document.apply_plan(plan)
    assert [block.text for block in updated.blocks] == [block.text for block in document.blocks]
    for before, after in zip(document.blocks, updated.blocks):
        if before.kind == "heading":
            assert all(run.format.color == "0000FF" for run in after.runs)
            for run in updated.paragraphs[after.id].findall(W + "r"):
                assert run.find(f"{W}rPr/{W}color").get(W + "val") == "0000FF"
        else:
            assert etree.tostring(document.paragraphs[before.id]) == etree.tostring(updated.paragraphs[after.id])
    with ZipFile(BytesIO(document.data)) as before, ZipFile(BytesIO(updated.data)) as after:
        for name in before.namelist():
            if name != MAIN_PART:
                assert before.read(name) == after.read(name)


def test_font_size_bold_italic_alignment_and_text_can_be_combined(document):
    """Combine text replacement with point-size, weight, italic, and alignment changes."""
    block = next(block for block in document.blocks if "30 days" in block.text)
    plan = EditPlan(
        summary="Format terms and edit text.",
        edits=[Edit(block_id=block.id, original="30", replacement="45")],
        formats=[FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(
            font_size_pt=18.5, bold=False, italic=True, alignment="center",
        ))],
    )
    updated = document.apply_plan(plan)
    result = next(b for b in updated.blocks if b.id == block.id)
    assert "45 days" in result.text
    assert result.format.alignment == "center"
    assert all(run.format.font_size_pt == 18.5 and run.format.bold is False and run.format.italic for run in result.runs)
    assert all(node.get(W + "val") == "37" for node in updated.paragraphs[block.id].iter(W + "sz"))


def test_font_family_is_written_to_all_scripts_and_survives_preview(document):
    """Replace theme fonts for every script and preserve the font choice in preview output."""
    block = document.blocks[0]
    def mutate(tree):
        """Add theme font attributes to a run before testing direct font replacement."""
        run = tree.find(W + "body").find(W + "p").find(W + "r")
        properties = run.find(W + "rPr")
        if properties is None:
            properties = etree.Element(W + "rPr")
            run.insert(0, properties)
        fonts = etree.Element(W + "rFonts", {
            W + "asciiTheme": "majorHAnsi", W + "hAnsiTheme": "majorHAnsi",
            W + "eastAsiaTheme": "majorEastAsia", W + "cstheme": "majorBidi",
        })
        properties.insert(0, fonts)
    source = WordDocument(change_xml(document.data, mutate), document.name)
    plan = EditPlan(summary="Use Libertinus Serif and blue headings", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text,
                   format=ParagraphFormat(font_family="Libertinus Serif", color="0000FF", font_size_pt=18)),
    ])
    changes = source.preview_plan(plan)
    assert "not embedded" in changes[0].description
    result = source.apply_plan(plan)
    assert all(run.format.font_family == "Libertinus Serif" for run in result.blocks[0].runs)
    for run in result.paragraphs[block.id].findall(W + "r"):
        fonts = run.find(f"{W}rPr/{W}rFonts")
        assert dict(fonts.attrib) == {W + script: "Libertinus Serif" for script in ("ascii", "hAnsi", "eastAsia", "cs")}
        tags = [etree.QName(child).localname for child in run.find(W + "rPr")]
        assert tags.index("rFonts") < tags.index("color") < tags.index("sz")
    assert [b.text for b in source.blocks] == [b.text for b in result.blocks]
    assert all(result.parts[name] == source.parts[name] for name in source.parts if name != MAIN_PART)


@pytest.mark.parametrize("family", ["", " " * 3, "a" * 81, "Arial; color:red", "url(http://example.com)", "Arial\n"])
def test_invalid_font_names_are_rejected_atomically(document, family):
    """Reject blank, oversized, and markup-like font family values without mutation."""
    block = document.blocks[0]
    with pytest.raises(DemoError, match="Font families"):
        document.apply_plan(EditPlan(summary="Invalid font", edits=[], formats=[
            FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(font_family=family)),
        ]))


def test_direct_color_overrides_theme_color(document):
    """Ensure a direct requested color replaces inherited theme color attributes."""
    block = next(block for block in document.blocks if block.kind == "heading")
    def mutate(tree):
        """Attach a theme-derived red run color to the selected heading."""
        p = tree.find(W + "body").find(W + "p")
        run = p.find(W + "r")
        properties = run.find(W + "rPr")
        if properties is None:
            properties = etree.Element(W + "rPr")
            run.insert(0, properties)
        etree.SubElement(properties, W + "color", {
            W + "val": "FF0000", W + "themeColor": "accent1", W + "themeTint": "99",
        })
    source = WordDocument(change_xml(document.data, mutate), document.name)
    updated = source.apply_plan(EditPlan(summary="Blue", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(color="0000ff")),
    ]))
    color = updated.paragraphs[block.id].find(f"{W}r/{W}rPr/{W}color")
    assert dict(color.attrib) == {W + "val": "0000FF"}


def test_real_table_has_dual_widths_headers_and_preserves_surrounding_content(document):
    """Verify generated table geometry, header metadata, content mapping, and package preservation."""
    table = scope_table(document)
    updated = document.apply_plan(EditPlan(
        summary="Blue headings and scope table.", edits=[], formats=heading_plan(document).formats, tables=[table],
    ))
    tables = updated.body.findall(W + "tbl")
    assert len(tables) == 2  # Existing milestone table plus new scope table.
    new = tables[-1]
    assert new.find(f"{W}tr/{W}trPr/{W}tblHeader") is not None
    assert [[
        "".join(cell.itertext()) for cell in row.findall(W + "tc")
    ] for row in new.findall(W + "tr")] == table.rows
    width = int(new.find(f"{W}tblPr/{W}tblW").get(W + "w"))
    assert sum(int(column.get(W + "w")) for column in new.find(W + "tblGrid")) == width
    assert all(int(cell.get(W + "w")) == width for cell in new.iter(W + "tcW"))
    assert all(new.find(f"{W}tblPr/{W}tblCellMar/{W}{side}") is not None for side in ("top", "left", "bottom", "right"))
    assert not list(new.iter(W + "numPr"))
    assert updated.parts["word/numbering.xml"] == document.parts["word/numbering.xml"]
    for name in document.parts:
        if name != MAIN_PART:
            assert updated.parts[name] == document.parts[name]
    nodes = updated.nodes()
    assert sum(node.kind == "table" for node in nodes) == 2
    assert [cell.blocks[0].text for row in nodes[-1].rows for cell in row] == [row[0] for row in table.rows]
    assert all(block.can_convert is False for block in updated.blocks if block.kind == "table")


def test_table_cells_escape_xml_and_support_multiple_columns(document):
    """Convert list blocks to a two-column table while safely escaping cell text."""
    blocks = [b for b in document.blocks if b.kind == "list"]
    originals = ["Alpha & <beta> | Week 1", "Gamma | Week 2"]
    def mutate(tree):
        """Replace list paragraph text with values containing XML-sensitive characters."""
        paragraphs = list(tree.find(W + "body").iter(W + "p"))
        for block, text in zip(blocks, originals):
            nodes = list(paragraphs[int(block.id[1:])].iter(W + "t"))
            nodes[0].text = text
            for node in nodes[1:]:
                node.text = ""
    source = WordDocument(change_xml(document.data, mutate), document.name)
    table = TableEdit(
        block_ids=[block.id for block in blocks], originals=originals,
        rows=[["Item", "Target"], ["Alpha & <beta>", "Week 1"], ["Gamma", "Week 2"]], header=True,
    )
    updated = source.apply_plan(EditPlan(summary="Table", edits=[], tables=[table]))
    new = updated.body.findall(W + "tbl")[-1]
    widths = [int(column.get(W + "w")) for column in new.find(W + "tblGrid")]
    assert len(widths) == 2
    for row in new.findall(W + "tr"):
        assert [int(cell.find(f"{W}tcPr/{W}tcW").get(W + "w")) for cell in row.findall(W + "tc")] == widths
    assert any(block.text == "Alpha & <beta>" for block in updated.blocks)
    assert b"&amp;" in updated.parts[MAIN_PART] and b"&lt;beta&gt;" in updated.parts[MAIN_PART]


@pytest.mark.parametrize("values", [
    {"color": "#0000FF"}, {"color": "red"}, {"color": "0000FF;display:none"},
    {"font_size_pt": 5}, {"font_size_pt": 72.5}, {"font_size_pt": 12.25},
    {"font_size_pt": float("nan")}, {"font_size_pt": float("inf")}, {},
])
def test_invalid_formatting_is_rejected_atomically(document, values):
    """Reject invalid color, point-size, or empty formatting and preserve source bytes."""
    block = document.blocks[0]
    before = document.data
    plan = EditPlan(summary="Invalid", edits=[], formats=[
        FormatEdit(block_id=block.id, original=block.text, format=ParagraphFormat(**values)),
    ])
    with pytest.raises(DemoError):
        document.apply_plan(plan)
    assert document.data == before


@pytest.mark.parametrize("case", ["missing", "stale", "duplicate", "readonly"])
def test_format_targets_must_be_current_unique_and_editable(document, case):
    """Require formatting targets to exist, match current text, be unique, and be editable."""
    block = document.blocks[0]
    if case == "readonly":
        def mutate(tree):
            """Insert a bookmark marker so the target paragraph becomes read-only."""
            p = tree.find(W + "body").find(W + "p")
            etree.SubElement(p, W + "bookmarkStart", {W + "id": "1", W + "name": "protected"})
        document = WordDocument(change_xml(document.data, mutate), document.name)
    edit = FormatEdit(block_id="missing" if case == "missing" else block.id,
                      original="stale" if case == "stale" else block.text, format=ParagraphFormat(bold=True))
    with pytest.raises(DemoError):
        document.apply_plan(EditPlan(summary="Invalid", edits=[], formats=[edit, edit] if case == "duplicate" else [edit]))


@pytest.mark.parametrize("case", [
    "loss", "reorder", "nonconsecutive", "duplicate", "stale", "missing", "existing_table",
    "ragged", "empty", "no_data", "overlap_format", "overlap_text", "duplicate_table", "bad_cell",
])
def test_unsafe_table_conversion_is_rejected_atomically(document, case):
    """Reject lossy, stale, malformed, overlapping, or otherwise unsafe table plans."""
    table = scope_table(document)
    formats, edits, tables = [], [], [table]
    if case == "loss":
        table.rows[-1][0] = "A fact was dropped."
    elif case == "reorder":
        table.block_ids.reverse()
        table.originals.reverse()
        table.rows[1:] = list(reversed(table.rows[1:]))
    elif case == "nonconsecutive":
        block = document.blocks[0]
        table.block_ids.insert(0, block.id)
        table.originals.insert(0, block.text)
        table.rows.insert(1, [block.text])
    elif case == "duplicate":
        table.block_ids[1] = table.block_ids[0]
    elif case == "stale":
        table.originals[0] = "Not current."
    elif case == "missing":
        table.block_ids[0] = "unknown"
    elif case == "existing_table":
        block = next(b for b in document.blocks if b.kind == "table")
        table.block_ids[0], table.originals[0] = block.id, block.text
    elif case == "ragged":
        table.rows[-1].append("extra cell")
    elif case == "empty":
        table.rows = []
    elif case == "no_data":
        table.rows = table.rows[:1]
    elif case == "overlap_format":
        formats = [FormatEdit(block_id=table.block_ids[0], original=table.originals[0], format=ParagraphFormat(bold=True))]
    elif case == "overlap_text":
        edits = [Edit(block_id=table.block_ids[0], original=table.originals[0], replacement="Changed")]
    elif case == "duplicate_table":
        tables.append(table)
    elif case == "bad_cell":
        table.rows[0][0] = "bad\nheader"
    before = document.data
    with pytest.raises(DemoError):
        document.apply_plan(EditPlan(summary="Invalid", edits=edits, formats=formats, tables=tables))
    assert document.data == before


def test_run_preview_resolves_existing_mixed_formatting(document):
    """Resolve existing bold, italic, and heading-size values in rich run previews."""
    block = next(b for b in document.blocks if "30 days" in b.text)
    assert any(run.format.bold and run.text == "30" for run in block.runs)
    assert any(run.format.italic and "$25,000" in run.text for run in block.runs)
    assert all(b.format.font_size_pt == 16 for b in document.blocks if b.kind == "heading")
