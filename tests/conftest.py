from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
import pytest

from app.documents import MAIN_PART, WordDocument

SAMPLE = Path(__file__).resolve().parent.parent / "demo" / "sample.docx"


@pytest.fixture
def sample_bytes():
    return SAMPLE.read_bytes()


@pytest.fixture
def document(sample_bytes):
    return WordDocument(sample_bytes, "sample.docx")


def repack(data, replacements):
    output = BytesIO()
    with ZipFile(BytesIO(data)) as original, ZipFile(output, "w") as updated:
        for entry in original.infolist():
            updated.writestr(entry, replacements.get(entry.filename, original.read(entry)))
        for name, value in replacements.items():
            if name not in original.namelist():
                updated.writestr(name, value)
    return output.getvalue()


def change_xml(data, mutate, part=MAIN_PART):
    with ZipFile(BytesIO(data)) as package:
        tree = etree.fromstring(package.read(part))
    mutate(tree)
    return repack(data, {part: etree.tostring(tree, encoding="UTF-8", xml_declaration=True)})
