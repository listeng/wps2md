"""wps2md — parse legacy WPS (.wps, OLE2 Word-binary) files into Markdown.

Quick start:

    from wps2md import parse, to_markdown

    doc = parse("file.wps")
    print(doc.main_text)
    print(to_markdown(doc.paragraphs))
"""
from wps2md.core import (
    Paragraph,
    WpsDocument,
    WpsParseError,
    parse,
    to_markdown,
)

__all__ = [
    "Paragraph",
    "WpsDocument",
    "WpsParseError",
    "parse",
    "to_markdown",
]
__version__ = "0.2.0"
