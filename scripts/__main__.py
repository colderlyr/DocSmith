"""DocSmith CLI — MD+LaTeX → professional .docx.

Usage:
    python3 -m scripts --output out.docx --config config.json --content content.md
    (run from docsmith/ directory)
"""

import json
import argparse

from .pipeline import TwoPassBuilder


def main():
    p = argparse.ArgumentParser(description="DocSmith — MD+LaTeX → professional .docx")
    p.add_argument("--output", required=True, help="Output .docx path")
    p.add_argument("--config", required=True, help="JSON config file")
    p.add_argument("--content", required=True, help="Markdown content file")
    args = p.parse_args()

    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)
    with open(args.content, 'r', encoding='utf-8') as f:
        md_text = f.read()

    builder = TwoPassBuilder(config, md_text)
    output_path = builder.build(args.output)
    print(f"OK — {output_path}")


if __name__ == "__main__":
    main()
