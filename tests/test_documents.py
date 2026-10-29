from io import BytesIO
import binascii
import random
import struct
import zlib
from zipfile import ZipFile

from lxml import etree
import pytest

from app import documents
from app.documents import Edit, MAIN_PART, W, WordDocument
from app.errors import DemoError
from conftest import change_xml, repack


def test_fixture_image_integrity(sample_bytes):
    with ZipFile(BytesIO(sample_bytes)) as package:
        image = package.read(next(name for name in package.namelist() if name.endswith(".png")))
    assert image[:8] == b"\x89PNG\r\n\x1a\n"
    offset = 8
    while offset < len(image):
        length = struct.unpack(">I", image[offset:offset + 4])[0]
        kind = image[offset + 4:offset + 8]
        data = image[offset + 8:offset + 8 + length]
        crc = struct.unpack(">I", image[offset + 8 + length:offset + 12 + length])[0]
        assert binascii.crc32(kind + data) & 0xFFFFFFFF == crc
        if kind == b"IDAT":
            assert zlib.decompress(data)
        offset += length + 12


def find_block(document, text):
    return next(block for block in document.blocks if text in block.text)


def test_round_trip_preserves_parts_and_run_formatting(document, sample_bytes):
    block = find_block(document, "30 days")
    updated = document.apply([Edit(block_id=block.id, original="30", replacement="45")])
    assert find_block(updated, "45 days").text == block.text.replace("30", "45")
    original_runs = list(document.paragraphs[block.id].iter(W + "r"))
    new_runs = list(updated.paragraphs[block.id].iter(W + "r"))
    assert len(original_runs) == len(new_runs)
    for before, after in zip(original_runs, new_runs):
        assert etree.tostring(before.find(W + "rPr")) == etree.tostring(after.find(W + "rPr")) if before.find(W + "rPr") is not None else after.find(W + "rPr") is None
    with ZipFile(BytesIO(sample_bytes)) as before, ZipFile(BytesIO(updated.data)) as after:
        assert set(before.namelist()) == set(after.namelist())
        assert any("media/" in name for name in before.namelist())
        for name in before.namelist():
            if name != MAIN_PART:
                assert before.read(name) == after.read(name), name
    assert document.data == sample_bytes
    assert updated.version != document.version


def test_cross_run_replacement_and_multiple_edits(document):
    block = find_block(document, "30 days")
    edits = [
        Edit(block_id=block.id, original="within 30 days", replacement="within 45 business days"),
        Edit(block_id=block.id, original="$25,000", replacement="$30,000"),
    ]
    updated = document.apply(edits)
    assert updated.blocks[[b.id for b in updated.blocks].index(block.id)].text == block.text.replace(
        "within 30 days", "within 45 business days"
    ).replace("$25,000", "$30,000")


@pytest.mark.parametrize("old,new,kind", [
    ("Week 4", "Week 6", "table"),
    ("Commercial terms", "Payment terms", "heading"),
    ("Create a reusable weekly status template.", "Build a weekly status template.", "list"),
    ("30", "45 & <60> \u201cillustrative\u201d", "paragraph"),
])
def test_supported_regions_and_xml_escaping(document, old, new, kind):
    block = find_block(document, old)
    assert block.kind == kind
    updated = document.apply([Edit(block_id=block.id, original=old, replacement=new)])
    assert new in find_block(updated, new).text
    etree.fromstring(updated.parts[MAIN_PART])


@pytest.mark.parametrize("replacement", ["bad\nline", "bad\ttab", "bad\x00character"])
def test_invalid_replacement_is_atomic(document, replacement):
    before = document.data
    block = find_block(document, "30 days")
    with pytest.raises(DemoError, match="unsupported"):
        document.apply([Edit(block_id=block.id, original="30", replacement=replacement)])
    assert document.data == before


def test_overlap_unknown_and_nonmatching_proposals(document):
    block = find_block(document, "30 days")
    with pytest.raises(DemoError, match="overlapping"):
        document.preview([
            Edit(block_id=block.id, original="30 days", replacement="45 days"),
            Edit(block_id=block.id, original="days of", replacement="business days of"),
        ])
    with pytest.raises(DemoError, match="unknown"):
        document.preview([Edit(block_id="invented", original="30", replacement="45")])
    for original in ("", "not present"):
        with pytest.raises(DemoError, match="uniquely"):
            document.preview([Edit(block_id=block.id, original=original, replacement="45")])


def test_ambiguous_source_text(document):
    block = find_block(document, "30 days")
    modified = document.apply([Edit(block_id=block.id, original="30 days", replacement="30 days and 30 days")])
    with pytest.raises(DemoError, match="uniquely"):
        modified.preview([Edit(block_id=block.id, original="30", replacement="45")])
    modified = document.apply([Edit(block_id=block.id, original="30 days", replacement="aaaa")])
    with pytest.raises(DemoError, match="uniquely"):
        modified.preview([Edit(block_id=block.id, original="aaa", replacement="b")])


def test_comment_or_bookmark_region_is_readonly(sample_bytes):
    def mutate(tree):
        paragraph = next(p for p in tree.iter(W + "p") if "30" in "".join(p.itertext()))
        paragraph.insert(1, etree.Element(W + "bookmarkStart", {W + "id": "1", W + "name": "terms"}))
    document = WordDocument(change_xml(sample_bytes, mutate), "sample.docx")
    block = find_block(document, "30 days")
    assert not block.editable
    assert any("read-only" in warning for warning in document.warnings)
    assert block not in document.ai_blocks()
    with pytest.raises(DemoError, match="read-only"):
        document.apply([Edit(block_id=block.id, original="30", replacement="45")])


def test_document_protection_and_macros_rejected(sample_bytes):
    def protect(tree):
        etree.SubElement(tree, W + "documentProtection", {W + "enforcement": "1"})
    with pytest.raises(DemoError, match="Protected"):
        WordDocument(change_xml(sample_bytes, protect, "word/settings.xml"), "sample.docx")
    with pytest.raises(DemoError, match="macro"):
        WordDocument(repack(sample_bytes, {"word/vbaProject.bin": b"synthetic"}), "sample.docx")


def test_leaf_revision_markers_are_readonly(sample_bytes):
    def mutate(tree):
        paragraph = next(p for p in tree.iter(W + "p") if "30" in "".join(p.itertext()))
        properties = paragraph.find(W + "pPr")
        if properties is None:
            properties = etree.Element(W + "pPr")
            paragraph.insert(0, properties)
        run_properties = etree.SubElement(properties, W + "rPr")
        etree.SubElement(run_properties, W + "del", {W + "id": "1", W + "author": "Synthetic"})
    document = WordDocument(change_xml(sample_bytes, mutate), "sample.docx")
    assert not find_block(document, "30 days").editable


def test_file_and_expansion_limits(sample_bytes, monkeypatch):
    with pytest.raises(DemoError, match="10 MiB"):
        WordDocument(b"x" * (10 * 1024 * 1024 + 1), "sample.docx")
    monkeypatch.setattr(documents, "MAX_EXPANDED_BYTES", 1)
    with pytest.raises(DemoError, match="expanded"):
        WordDocument(sample_bytes, "sample.docx")


def test_ai_character_limit_is_not_silent_truncation(document, monkeypatch):
    monkeypatch.setattr(documents, "MAX_AI_CHARACTERS", 10)
    with pytest.raises(DemoError, match="30,000"):
        document.ai_blocks()


@pytest.mark.parametrize("data,name", [(b"not a zip", "file.docx"), (b"", "file.doc")])
def test_bad_files(data, name):
    with pytest.raises(DemoError):
        WordDocument(data, name)


def test_rights_managed_container_has_specific_error():
    data = bytes.fromhex("d0cf11e0a1b11ae1") + b"\x00" * 40 + "EncryptedPackage".encode("utf-16le")
    with pytest.raises(DemoError, match="still encrypted or rights-managed"):
        WordDocument(data, "test.docx")


def test_legacy_container_is_not_mislabeled_as_encrypted():
    with pytest.raises(DemoError, match="older Word or protected Office container"):
        WordDocument(bytes.fromhex("d0cf11e0a1b11ae1") + b"\x00" * 40, "test.docx")


def test_broken_zip_is_not_mislabeled_as_encrypted():
    with pytest.raises(DemoError, match="unreadable or damaged"):
        WordDocument(b"PK\x03\x04broken", "test.docx")


def test_doctype_is_rejected(sample_bytes):
    original = WordDocument(sample_bytes, "sample.docx")
    main = original.parts[MAIN_PART]
    split = main.find(b"?>")
    main = main[:split + 2] + b'<!DOCTYPE document [<!ENTITY x "untrusted">]>' + main[split + 2:]
    with pytest.raises(DemoError, match="document-type"):
        WordDocument(repack(sample_bytes, {MAIN_PART: main}), "sample.docx")


def test_randomized_span_patches_keep_exact_text(document):
    rng = random.Random(127)
    block = find_block(document, "30 days")
    for _ in range(80):
        start = rng.randrange(len(block.text))
        end = rng.randrange(start + 1, len(block.text) + 1)
        original = block.text[start:end]
        if block.text.count(original) != 1:
            continue
        replacement = "".join(rng.choices("Abc 145<&", k=rng.randrange(0, 40)))
        updated = document.apply([Edit(block_id=block.id, original=original, replacement=replacement)])
        text = next(b.text for b in updated.blocks if b.id == block.id)
        assert text == block.text[:start] + replacement + block.text[end:]
