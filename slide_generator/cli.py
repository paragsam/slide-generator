"""Command-line interface for slide-generator."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .generator import SlideGenerator


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="slide-generator",
        description=(
            "Fill a PowerPoint template's tables with data and produce a "
            "presentation-ready .pptx deck."
        ),
    )
    parser.add_argument("template", help="Path to the .pptx template file")
    parser.add_argument("data", help="Path to a JSON file containing fill data")
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="Output .pptx path (default: <template-stem>.out.pptx in ./output/)",
    )

    args = parser.parse_args(argv)

    template_path = Path(args.template)
    data_path = Path(args.data)

    if not data_path.exists():
        sys.exit(f"Error: data file not found: {data_path}")

    with data_path.open() as f:
        data = json.load(f)

    output_path = (
        Path(args.output)
        if args.output
        else Path("output") / f"{template_path.stem}.out.pptx"
    )

    gen = SlideGenerator(template_path)
    result = gen.generate(data, output_path)
    print(f"Saved: {result}")


if __name__ == "__main__":
    main()
