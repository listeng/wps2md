"""CLI: wps2md <file.wps>  (or: python -m wps2md <file.wps>)"""
import sys

from wps2md.core import WpsParseError, parse, to_markdown


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: wps2md <file.wps>", file=sys.stderr)
        return 2
    try:
        doc = parse(sys.argv[1])
    except WpsParseError as e:
        print(f"parse error: {e}", file=sys.stderr)
        return 1
    print(to_markdown(doc.paragraphs), end="")
    return 0


if __name__ == "__main__":
    sys.exit(main())
