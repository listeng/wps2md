# wps2md

A tiny Python library and CLI for converting legacy **WPS Writer `.wps`**
and **Word 97-2003 `.doc`** files (OLE2 Word-binary format, FIB magic
`0xA5EC`/`0xA5DC`) into structured text and Markdown.

Unlike `.docx` (which is OOXML/zip and can be read by `python-docx`),
`.wps` files saved by WPS Office are binary OLE2 compound documents.
This library reads the `WordDocument` stream, validates the FIB,
recovers paragraph style indices (`istd`) via `PlcfBtePapx` → FKPs,
and renders Heading 1-9 styles as `#`..`#########` in Markdown.

## Install

```bash
pip install wps2md
```

## CLI

```bash
wps2md example.wps                 # print Markdown to stdout
wps2md example.wps > example.md
wps2md example.doc                 # .doc files also supported
python -m wps2md example.wps       # equivalent
```

## Library

```python
from wps2md import parse, to_markdown

doc = parse("example.wps")
print(doc.main_text)                # plain text of the main body
print(doc.num_pages)                # from OLE SummaryInformation
print(to_markdown(doc.paragraphs))  # Markdown with H1-H9 from Word styles

for p in doc.paragraphs:
    print(p.heading_level, p.text)  # 0 for normal text, 1-9 for headings
```

## API

- `parse(path) -> WpsDocument` — parse a `.wps` or `.doc` file.
- `WpsDocument` — dataclass with `main_text`, `paragraphs`, `footnotes`,
  `headers_footers`, `annotations`, `encoding`, `num_pages`.
- `Paragraph(istd: int, text: str)` — one paragraph; `heading_level`
  returns 1-9 for built-in Heading styles, else 0.
- `to_markdown(paragraphs) -> str` — render paragraphs as Markdown.
- `WpsParseError` — raised for unsupported extensions, encrypted files,
  or unreadable streams.

## Limitations

- Tables, images, footnotes/headers paragraph styles, complex fields,
  and CHPX (character formatting like bold/italic) are not currently
  surfaced — only paragraph-level Heading styles drive Markdown output.
- Encrypted/password-protected files are rejected.
- Only the OLE2 Word-binary variant of `.wps` is supported (modern WPS
  Office still writes this for `.wps`; the OOXML `.docx` variant should
  be read with `python-docx` instead).

## License

MIT
