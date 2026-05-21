import pathlib

import pytest

from wps2md import WpsParseError, parse, to_markdown

SAMPLE = pathlib.Path(__file__).parent / "sample.wps"


@pytest.fixture(scope="module")
def doc():
    if not SAMPLE.exists():
        pytest.skip("sample.wps not present")
    return parse(SAMPLE)


def test_main_text_nonempty(doc):
    assert len(doc.main_text) > 0


def test_paragraphs_have_first_heading(doc):
    assert doc.paragraphs, "expected at least one paragraph"
    assert doc.paragraphs[0].heading_level == 1


def test_markdown_starts_with_h1(doc):
    md = to_markdown(doc.paragraphs)
    assert md.startswith("# ")


def test_only_h1_in_sample(doc):
    md = to_markdown(doc.paragraphs)
    h_lines = [ln for ln in md.splitlines() if ln.startswith("#")]
    assert all(ln.startswith("# ") for ln in h_lines)


def test_rejects_unsupported_extension(tmp_path):
    p = tmp_path / "foo.txt"
    p.write_bytes(b"\xd0\xcf\x11\xe0")
    with pytest.raises(WpsParseError):
        parse(p)
