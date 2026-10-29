from collections import defaultdict
from difflib import SequenceMatcher
from hashlib import sha256
from io import BytesIO
import re
from typing import Literal
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile, ZipInfo

from lxml import etree
from pydantic import BaseModel, ConfigDict

from app.errors import DemoError

MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_EXPANDED_BYTES = 50 * 1024 * 1024
MAX_AI_CHARACTERS = 30_000
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MAIN_PART = "word/document.xml"
DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
INVALID_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")


class Block(BaseModel):
    id: str
    text: str
    kind: Literal["paragraph", "heading", "table", "list"]
    editable: bool
    reason: str | None = None
    format: "ParagraphFormat"
    runs: list["PreviewRun"]
    can_convert: bool = False


class ParagraphFormat(BaseModel):
    model_config = ConfigDict(extra="forbid")
    color: str | None = None
    font_size_pt: float | None = None
    font_family: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    alignment: Literal["left", "center", "right", "justify"] | None = None


class PreviewRun(BaseModel):
    text: str
    format: ParagraphFormat


class PreviewCell(BaseModel):
    blocks: list[Block]
    colspan: int = 1


class PreviewNode(BaseModel):
    kind: Literal["paragraph", "table"]
    blocks: list[Block] = []
    rows: list[list[PreviewCell]] = []


class FormatEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    block_id: str
    original: str
    format: ParagraphFormat


class TableEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    block_ids: list[str]
    originals: list[str]
    rows: list[list[str]]
    header: bool


class Edit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    block_id: str
    original: str
    replacement: str


class EditPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")
    summary: str
    edits: list[Edit]
    formats: list[FormatEdit] = []
    tables: list[TableEdit] = []


class Change(BaseModel):
    block_id: str
    before: str
    after: str
    kind: Literal["text", "format", "table"] = "text"
    description: str = ""


def _properties(parent: etree._Element, name: str) -> etree._Element:
    found = parent.find(W + name)
    if found is None:
        found = etree.Element(W + name)
        parent.insert(0, found)
    return found


def _set_property(parent: etree._Element, name: str, value: str, before: tuple[str, ...]) -> None:
    node = parent.find(W + name)
    if node is None:
        node = etree.Element(W + name)
        position = next((i for i, child in enumerate(parent) if child.tag in {W + tag for tag in before}), len(parent))
        parent.insert(position, node)
    node.set(W + "val", value)


def _valid_text(text: str) -> bool:
    return not INVALID_XML.search(text) and not any(char in text for char in "\r\n\t")


def _format_values(properties: etree._Element | None) -> dict[str, object]:
    values: dict[str, object] = {}
    if properties is None:
        return values
    fonts = properties.find(W + "rFonts")
    if fonts is not None:
        family = fonts.get(W + "ascii") or fonts.get(W + "hAnsi")
        if family and re.fullmatch(r"[\w][\w .()'-]{0,79}", family):
            values["font_family"] = family
    for prop, field in (("b", "bold"), ("i", "italic")):
        node = properties.find(W + prop)
        if node is not None:
            values[field] = node.get(W + "val", "1") not in ("0", "false", "off")
    color = properties.find(W + "color")
    if color is not None and re.fullmatch(r"[0-9A-Fa-f]{6}", color.get(W + "val", "")):
        values["color"] = color.get(W + "val")
    size = properties.find(W + "sz")
    if size is not None:
        try:
            points = float(size.get(W + "val", "")) / 2
            if 1 <= points <= 400:
                values["font_size_pt"] = points
        except ValueError:
            pass  # Some Word producers emit a nonnumeric size; leave it inherited.
    return values


def _xml(data: bytes) -> etree._Element:
    try:
        tree = etree.fromstring(data, parser=etree.XMLParser(
            resolve_entities=False, no_network=True, load_dtd=False,
        ))
    except etree.XMLSyntaxError as exc:
        raise DemoError("The Word package contains malformed XML.") from exc
    if tree.getroottree().docinfo.doctype:
        raise DemoError("Documents with XML document-type declarations are not supported.")
    return tree


def _text_nodes(paragraph: etree._Element) -> list[etree._Element]:
    return [
        node for node in paragraph.iter(W + "t")
        if next(node.iterancestors(W + "p"), None) is paragraph
    ]


def _read_only_reason(paragraph: etree._Element) -> str | None:
    complex_ancestors = {W + name for name in (
        "sdt", "ins", "del", "txbxContent", "customXml", "moveFrom", "moveTo",
    )}
    if any(parent.tag in complex_ancestors for parent in paragraph.iterancestors()):
        return "Text in a content control, revision, or text box is read-only."
    for child in paragraph:
        if child.tag == W + "pPr":
            revision_tags = {W + name for name in ("rPrChange", "pPrChange", "del", "ins", "moveFrom", "moveTo")}
            if any(node.tag in revision_tags for node in child.iter()):
                return "This paragraph contains tracked formatting changes."
            continue
        if child.tag != W + "r":
            return "This paragraph contains links, comments, bookmarks, or complex content."
        if any(node.tag not in (W + "rPr", W + "t") for node in child):
            return "This paragraph contains a field, drawing, tab, or line break."
        if next(child.iter(W + "rPrChange"), None) is not None:
            return "This paragraph contains tracked formatting changes."
    return None


class WordDocument:
    def __init__(self, data: bytes, name: str) -> None:
        if not name.lower().endswith(".docx"):
            raise DemoError("Choose a .docx file. Legacy .doc and macro-enabled files are not supported.")
        if len(data) > MAX_FILE_BYTES:
            raise DemoError("This demo accepts Word files up to 10 MiB.", 413)
        if data.startswith(bytes.fromhex("d0cf11e0a1b11ae1")):
            if "EncryptedPackage".encode("utf-16le") in data or "DRMEncryptedDataSpace".encode("utf-16le") in data:
                raise DemoError(
                    "This downloaded file is still encrypted or rights-managed, not a readable DOCX package. "
                    "Removing the visible sensitivity label does not confirm that protection has been removed. "
                    "Use a synthetic, unprotected demo document permitted by your organization's policy, "
                    "or ask its owner/admin to confirm the protection and saved cloud version."
                )
            raise DemoError(
                "This file uses an older Word or protected Office container, not a standard DOCX ZIP package. "
                "Open it in Word and, if permitted, save a standard unprotected .docx for the demo."
            )
        self.name = name
        self.data = data
        self.version = sha256(data).hexdigest()
        self.parts: dict[str, bytes] = {}
        self.entries: list[ZipInfo] = []
        try:
            with ZipFile(BytesIO(data)) as package:
                entries = package.infolist()
                if len(entries) > 2000 or sum(item.file_size for item in entries) > MAX_EXPANDED_BYTES:
                    raise DemoError("The expanded Word package exceeds the demo's 50 MiB/2,000-part limit.", 413)
                names = [entry.filename for entry in entries]
                if len(set(names)) != len(names):
                    raise DemoError("The Word package contains duplicate parts.")
                for entry in entries:
                    path = entry.filename
                    if (
                        entry.flag_bits & 1 or path.startswith(("/", "\\"))
                        or "\\" in path or ".." in path.split("/")
                    ):
                        raise DemoError("Encrypted or unsafe Word packages are not supported.")
                    if path.lower().startswith("_xmlsignatures/") or path.lower().endswith("vbaproject.bin"):
                        raise DemoError("Signed or macro-enabled documents are not supported.")
                    self.parts[path] = package.read(entry)
                self.entries = entries
        except BadZipFile as exc:
            raise DemoError("The DOCX ZIP package is unreadable or damaged. Download a fresh copy from OneDrive and retry.") from exc
        except (RuntimeError, NotImplementedError) as exc:
            raise DemoError("The Word package uses unsupported ZIP encryption or compression.") from exc
        if MAIN_PART not in self.parts or "[Content_Types].xml" not in self.parts:
            raise DemoError("This file is missing required Word document parts.")
        types = _xml(self.parts["[Content_Types].xml"])
        main_types = [
            node.get("ContentType") for node in types
            if node.get("PartName") == "/" + MAIN_PART
        ]
        if main_types != [DOCX_TYPE + ".main+xml"]:
            raise DemoError("Only ordinary, non-macro .docx documents are supported.")
        if "word/settings.xml" in self.parts:
            settings = _xml(self.parts["word/settings.xml"])
            if (
                next(settings.iter(W + "documentProtection"), None) is not None
                or next(settings.iter(W + "writeProtection"), None) is not None
            ):
                raise DemoError("Protected documents are not editable in this demo.")
        self.tree = _xml(self.parts[MAIN_PART])
        if self.tree.tag != W + "document":
            raise DemoError("This Word XML format is not supported. Save it as a standard .docx in Word.")
        body = self.tree.find(W + "body")
        if body is None:
            raise DemoError("The Word document has no body.")
        self.blocks: list[Block] = []
        self.paragraphs: dict[str, etree._Element] = {}
        self.styles: dict[str, etree._Element] = {}
        self.default_format: dict[str, object] = {}
        self.default_style = ""
        if "word/styles.xml" in self.parts:
            styles = _xml(self.parts["word/styles.xml"])
            self.styles = {node.get(W + "styleId", ""): node for node in styles.iter(W + "style")}
            self.default_format = _format_values(styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr"))
            self.default_style = next((
                name for name, node in self.styles.items()
                if node.get(W + "type") == "paragraph" and node.get(W + "default") == "1"
            ), "")
        self.body = body
        for index, paragraph in enumerate(body.iter(W + "p")):
            nodes = _text_nodes(paragraph)
            text = "".join(node.text or "" for node in nodes)
            if not text:
                continue
            block_id = f"p{index}"
            reason = _read_only_reason(paragraph)
            style = paragraph.find(f"{W}pPr/{W}pStyle")
            style_name = style.get(W + "val", "") if style is not None else ""
            kind: Literal["paragraph", "heading", "table", "list"] = "paragraph"
            if next(paragraph.iterancestors(W + "tc"), None) is not None:
                kind = "table"
            elif style_name.lower().startswith(("heading", "title")):
                kind = "heading"
            elif paragraph.find(f"{W}pPr/{W}numPr") is not None or style_name.lower().startswith("list"):
                kind = "list"
            properties = paragraph.find(W + "pPr")
            values = {**self.default_format, **self._style_format(style_name or self.default_style)}
            if properties is not None:
                values.update(_format_values(properties.find(W + "rPr")))
                alignment = properties.find(W + "jc")
                if alignment is not None:
                    align = alignment.get(W + "val")
                    if align in ("left", "center", "right", "both"):
                        values["alignment"] = "justify" if align == "both" else align
            preview_format = ParagraphFormat.model_validate(values)
            runs = [
                PreviewRun(text=node.text or "", format=ParagraphFormat.model_validate({
                    **values, **self._run_format(node.getparent()),
                }))
                for node in nodes
            ]
            self.blocks.append(Block(
                id=block_id, text=text, kind=kind, editable=reason is None, reason=reason,
                format=preview_format, runs=runs,
                can_convert=reason is None and paragraph.getparent() is body
                and paragraph.find(f"{W}pPr/{W}sectPr") is None,
            ))
            self.paragraphs[block_id] = paragraph
        if not self.blocks:
            raise DemoError("No readable main-body text was found in this document.")
        self.warnings = [
            "Content preview only: Word pagination/layout is not reproduced. Headers, footers, and images are preserved, not previewed.",
        ]
        readonly = sum(not block.editable for block in self.blocks)
        if readonly:
            self.warnings.append(f"{readonly} complex text block(s) are read-only and excluded from AI editing.")
        if (
            next(self.tree.iter(W + "del"), None) is not None
            or next(self.tree.iter(W + "ins"), None) is not None
        ):
            self.warnings.append("Existing tracked changes are preserved; deleted text is not shown in this preview.")

    def _style_format(self, name: str, seen: frozenset[str] = frozenset()) -> dict[str, object]:
        if not name or name in seen or len(seen) >= 20 or name not in self.styles:
            return {}
        style = self.styles[name]
        base = style.find(W + "basedOn")
        values = self._style_format(base.get(W + "val", "") if base is not None else "", seen | {name})
        values.update(_format_values(style.find(W + "rPr")))
        alignment = style.find(f"{W}pPr/{W}jc")
        if alignment is not None and alignment.get(W + "val") in ("left", "center", "right", "both"):
            values["alignment"] = "justify" if alignment.get(W + "val") == "both" else alignment.get(W + "val")
        return values

    def _run_format(self, run: etree._Element | None) -> dict[str, object]:
        if run is None:
            return {}
        properties = run.find(W + "rPr")
        style = properties.find(W + "rStyle") if properties is not None else None
        return {
            **self._style_format(style.get(W + "val", "") if style is not None else ""),
            **_format_values(properties),
        }

    def nodes(self) -> list[PreviewNode]:
        by_element = {self.paragraphs[block.id]: block for block in self.blocks}
        result: list[PreviewNode] = []
        for child in self.body:
            if child.tag == W + "tbl":
                rows: list[list[PreviewCell]] = []
                for row in child.findall(W + "tr"):
                    cells: list[PreviewCell] = []
                    for cell in row.findall(W + "tc"):
                        span = cell.find(f"{W}tcPr/{W}gridSpan")
                        value = span.get(W + "val", "1") if span is not None else "1"
                        cells.append(PreviewCell(
                            blocks=[by_element[p] for p in cell.iter(W + "p") if p in by_element],
                            colspan=min(20, max(1, int(value))) if value.isdigit() else 1,
                        ))
                    rows.append(cells)
                result.append(PreviewNode(kind="table", rows=rows))
            else:
                paragraphs = [child] if child.tag == W + "p" else list(child.iter(W + "p"))
                for paragraph in paragraphs:
                    if paragraph in by_element:
                        result.append(PreviewNode(kind="paragraph", blocks=[by_element[paragraph]]))
        return result

    def ai_blocks(self) -> list[Block]:
        blocks = [block for block in self.blocks if block.editable]
        if not blocks:
            raise DemoError("There are no supported editable text blocks in this document.")
        if sum(len(block.text) for block in blocks) > MAX_AI_CHARACTERS:
            raise DemoError("This demo sends at most 30,000 editable text characters. Use a smaller document.", 413)
        return blocks

    def preview(self, edits: list[Edit]) -> list[Change]:
        blocks = {block.id: block for block in self.blocks}
        grouped: dict[str, list[tuple[int, int, str]]] = defaultdict(list)
        if len(edits) > 50:
            raise DemoError("A proposal can contain at most 50 edits.")
        for edit in edits:
            block = blocks.get(edit.block_id)
            if block is None or not block.editable:
                raise DemoError("The proposal targets an unknown or read-only text block.")
            start = block.text.find(edit.original)
            if not edit.original or start < 0 or block.text.find(edit.original, start + 1) != -1:
                raise DemoError("An edit does not uniquely match the original text. Ask for a more specific change.")
            if INVALID_XML.search(edit.replacement) or "\n" in edit.replacement or "\r" in edit.replacement or "\t" in edit.replacement:
                raise DemoError("The proposal contains unsupported characters or structural changes.")
            grouped[block.id].append((start, start + len(edit.original), edit.replacement))
        changes: list[Change] = []
        for block_id, spans in grouped.items():
            spans.sort()
            if any(left[1] > right[0] for left, right in zip(spans, spans[1:])):
                raise DemoError("The proposal contains overlapping edits. Generate a new proposal.")
            before = blocks[block_id].text
            after = before
            for start, end, replacement in reversed(spans):
                after = after[:start] + replacement + after[end:]
            if before != after:
                changes.append(Change(block_id=block_id, before=before, after=after))
        after_text = {change.block_id: change.after for change in changes}
        if sum(len(after_text.get(block.id, block.text)) for block in self.blocks if block.editable) > MAX_AI_CHARACTERS:
            raise DemoError("The proposed result exceeds the 30,000-character demo limit.", 413)
        return changes

    def preview_plan(self, plan: EditPlan) -> list[Change]:
        if len(plan.summary) > 2000 or len(plan.edits) + len(plan.formats) + len(plan.tables) > 50:
            raise DemoError("A proposal must have a summary under 2,000 characters and at most 50 operations.")
        changes = self.preview(plan.edits)
        blocks = {block.id: block for block in self.blocks}
        formatted: set[str] = set()
        converted: set[str] = set()
        for edit in plan.formats:
            block = blocks.get(edit.block_id)
            if block is None or not block.editable or edit.original != block.text:
                raise DemoError("A formatting edit targets unknown, read-only, or stale text.")
            if edit.block_id in formatted:
                raise DemoError("The proposal contains duplicate formatting operations.")
            formatted.add(edit.block_id)
            values = edit.format.model_dump(exclude_none=True)
            if not values:
                raise DemoError("A formatting operation must specify at least one formatting property.")
            color = edit.format.color
            size = edit.format.font_size_pt
            family = edit.format.font_family
            if family is not None and not re.fullmatch(r"[\w][\w .()'-]{0,79}", family):
                raise DemoError("Font families must be 1-80 characters using letters, numbers, spaces, periods, hyphens, apostrophes or parentheses.")
            if color is not None and not re.fullmatch(r"[0-9A-Fa-f]{6}", color):
                raise DemoError("Font colors must be six-digit RGB values, without a # prefix.")
            if size is not None and (not 6 <= size <= 72 or size * 2 != int(size * 2)):
                raise DemoError("Font sizes must be between 6 and 72 points, in half-point increments.")
            description = ", ".join(f"{key.replace('_', ' ')}: {value}" for key, value in values.items())
            if family is not None:
                description += ". Fonts are not embedded; Word and your browser may substitute if this font is not installed."
            changes.append(Change(
                block_id=block.id, before=block.text, after=block.text, kind="format",
                description=description,
            ))
        for edit in plan.tables:
            if not edit.block_ids or len(edit.block_ids) > 50 or len(set(edit.block_ids)) != len(edit.block_ids):
                raise DemoError("A table conversion needs 1-50 unique source paragraphs.")
            if len(edit.originals) != len(edit.block_ids):
                raise DemoError("Table conversion source text is missing.")
            paragraphs: list[etree._Element] = []
            for block_id, original in zip(edit.block_ids, edit.originals):
                block = blocks.get(block_id)
                if block is None or not block.can_convert or original != block.text:
                    raise DemoError("Only current, editable main-body paragraphs or lists can be converted to a table.")
                if block_id in converted or block_id in formatted or any(e.block_id == block_id for e in plan.edits):
                    raise DemoError("Table conversions cannot overlap text, formatting, or other table operations.")
                converted.add(block_id)
                paragraphs.append(self.paragraphs[block_id])
            if any(left.getnext() is not right for left, right in zip(paragraphs, paragraphs[1:])):
                raise DemoError("Select consecutive paragraphs in document order for a table conversion.")
            if not edit.rows or len(edit.rows) > 51 or not 1 <= len(edit.rows[0]) <= 8:
                raise DemoError("A table must contain 1-50 data rows and 1-8 columns.")
            width = len(edit.rows[0])
            if any(len(row) != width for row in edit.rows) or (edit.header and len(edit.rows) < 2):
                raise DemoError("Table rows must have the same number of columns and at least one data row.")
            if len(edit.rows) - int(edit.header) > 50:
                raise DemoError("A table can have at most 50 data rows.")
            if any(not _valid_text(cell) or len(cell) > MAX_AI_CHARACTERS for row in edit.rows for cell in row):
                raise DemoError("Table cells contain unsupported or oversized text.")
            data_rows = edit.rows[1:] if edit.header else edit.rows
            normalize = lambda text: re.sub(r"[\s|;]+", "", text)
            source = normalize(" ".join(edit.originals))
            target = normalize(" ".join(cell for row in data_rows for cell in row))
            if source != target:
                raise DemoError("Table conversion must preserve source text and its order. Only spaces, pipes, or semicolon separators may be removed; headers may be added.")
            total = sum(len(cell) for row in edit.rows for cell in row)
            if total > MAX_AI_CHARACTERS:
                raise DemoError("The proposed table exceeds the 30,000-character demo limit.", 413)
            changes.append(Change(
                block_id=", ".join(edit.block_ids), before="\n".join(edit.originals),
                after="\n".join(" | ".join(row) for row in edit.rows), kind="table",
                description=f"Convert to a real Word table: {len(edit.rows)} rows, {width} columns. "
                "Cell text uses the first source paragraph's base formatting; source list markers are removed.",
            ))
        return changes

    def _format_paragraph(self, paragraph: etree._Element, formatting: ParagraphFormat) -> None:
        if formatting.alignment is not None:
            properties = _properties(paragraph, "pPr")
            _set_property(properties, "jc", "both" if formatting.alignment == "justify" else formatting.alignment, (
                "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
            ))
        for run in paragraph.findall(W + "r"):
            properties = _properties(run, "rPr")
            if formatting.font_family is not None:
                fonts = properties.find(W + "rFonts")
                if fonts is None:
                    fonts = etree.Element(W + "rFonts")
                    style = properties.find(W + "rStyle")
                    properties.insert(1 if style is not None else 0, fonts)
                for script in ("ascii", "hAnsi", "eastAsia", "cs"):
                    fonts.set(W + script, formatting.font_family)
                    fonts.attrib.pop(W + script + "Theme", None)
                fonts.attrib.pop(W + "cstheme", None)
            if formatting.bold is not None:
                _set_property(properties, "b", "1" if formatting.bold else "0", (
                    "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
                    "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
                    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
                    "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
                    "specVanish", "oMath", "rPrChange",
                ))
            if formatting.italic is not None:
                _set_property(properties, "i", "1" if formatting.italic else "0", (
                    "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
                    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position",
                    "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
                    "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
                ))
            if formatting.color is not None:
                _set_property(properties, "color", formatting.color.upper(), (
                    "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr",
                    "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
                ))
                color = properties.find(W + "color")
                if color is not None:
                    for attr in ("themeColor", "themeTint", "themeShade"):
                        color.attrib.pop(W + attr, None)
            if formatting.font_size_pt is not None:
                for name in ("sz", "szCs"):
                    _set_property(properties, name, str(int(formatting.font_size_pt * 2)), (
                        *(("szCs",) if name == "sz" else ()), "highlight", "u", "effect", "bdr", "shd",
                        "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath", "rPrChange",
                    ))

    def _convert_table(self, edit: TableEdit) -> None:
        paragraphs = [self.paragraphs[block_id] for block_id in edit.block_ids]
        section = paragraphs[0].find(f"{W}pPr/{W}sectPr")
        if section is None:
            section = self.body.find(W + "sectPr")
        content_width = 9360
        if section is not None:
            page_size, margin = section.find(W + "pgSz"), section.find(W + "pgMar")
            if page_size is not None and margin is not None:
                try:
                    content_width = int(page_size.get(W + "w", "12240")) - int(margin.get(W + "left", "1440")) - int(margin.get(W + "right", "1440"))
                except ValueError as exc:
                    raise DemoError("The document has an unsupported page-width value.") from exc
        if not 1440 <= content_width <= 31680:
            raise DemoError("The document's page width is outside the supported table range.")
        count = len(edit.rows[0])
        widths = [content_width // count] * count
        widths[-1] += content_width - sum(widths)
        table = etree.Element(W + "tbl")
        properties = etree.SubElement(table, W + "tblPr")
        etree.SubElement(properties, W + "tblW", {W + "w": str(content_width), W + "type": "dxa"})
        borders = etree.SubElement(properties, W + "tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            etree.SubElement(borders, W + side, {W + "val": "single", W + "sz": "4", W + "color": "919191"})
        etree.SubElement(properties, W + "tblLayout", {W + "type": "fixed"})
        margins = etree.SubElement(properties, W + "tblCellMar")
        for side in ("top", "left", "bottom", "right"):
            etree.SubElement(margins, W + side, {W + "w": "100", W + "type": "dxa"})
        grid = etree.SubElement(table, W + "tblGrid")
        for width in widths:
            etree.SubElement(grid, W + "gridCol", {W + "w": str(width)})
        first_block = next(block for block in self.blocks if block.id == edit.block_ids[0])
        base_format = first_block.format.model_copy(update={"alignment": "left"})
        for index, row in enumerate(edit.rows):
            tr = etree.SubElement(table, W + "tr")
            if edit.header and index == 0:
                etree.SubElement(etree.SubElement(tr, W + "trPr"), W + "tblHeader")
            for text, width in zip(row, widths):
                cell = etree.SubElement(tr, W + "tc")
                cell_properties = etree.SubElement(cell, W + "tcPr")
                etree.SubElement(cell_properties, W + "tcW", {W + "w": str(width), W + "type": "dxa"})
                paragraph = etree.SubElement(cell, W + "p")
                run = etree.SubElement(paragraph, W + "r")
                etree.SubElement(run, W + "t", {XML_SPACE: "preserve"}).text = text
                self._format_paragraph(paragraph, base_format.model_copy(update={"bold": True}) if edit.header and index == 0 else base_format)
        paragraphs[0].addprevious(table)
        for paragraph in paragraphs:
            self.body.remove(paragraph)

    def apply_plan(self, plan: EditPlan) -> "WordDocument":
        changes = self.preview_plan(plan)
        if not changes:
            raise DemoError("This proposal contains no changes.")
        updated = WordDocument(self.data, self.name)
        for change in (change for change in changes if change.kind == "text"):
            nodes = _text_nodes(updated.paragraphs[change.block_id])
            positions: list[tuple[etree._Element, int, int]] = []
            offset = 0
            for node in nodes:
                length = len(node.text or "")
                positions.append((node, offset, offset + length))
                offset += length
            differences = SequenceMatcher(None, change.before, change.after).get_opcodes()
            for tag, start, end, new_start, new_end in reversed(differences):
                if tag == "equal":
                    continue
                replacement = change.after[new_start:new_end]
                if start == end:
                    target = next(
                        ((node, left) for node, left, right in positions if left <= start <= right),
                        None,
                    )
                    if target is None:
                        raise DemoError("Unable to place the proposed insertion safely.")
                    node, left = target
                    local = start - left
                    node.text = (node.text or "")[:local] + replacement + (node.text or "")[local:]
                    node.set(XML_SPACE, "preserve")
                    continue
                inserted = False
                for node, left, right in positions:
                    if left >= end or right <= start:
                        continue
                    text = node.text or ""
                    local_start = max(start - left, 0)
                    local_end = min(end - left, right - left)
                    node.text = text[:local_start] + (replacement if not inserted else "") + text[local_end:]
                    node.set(XML_SPACE, "preserve")
                    inserted = True
        for edit in plan.formats:
            updated._format_paragraph(updated.paragraphs[edit.block_id], edit.format)
        for edit in plan.tables:
            updated._convert_table(edit)
        output = BytesIO()
        main_xml = etree.tostring(updated.tree, encoding="UTF-8", xml_declaration=True, standalone=True)
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as package:
            for entry in updated.entries:
                package.writestr(entry, main_xml if entry.filename == MAIN_PART else updated.parts[entry.filename])
        result = WordDocument(output.getvalue(), self.name)
        if sum(len(block.text) for block in result.blocks if block.editable) > MAX_AI_CHARACTERS:
            raise DemoError("The proposed result exceeds the 30,000-character demo limit.", 413)
        return result

    def apply(self, edits: list[Edit]) -> "WordDocument":
        return self.apply_plan(EditPlan(summary="", edits=edits))
