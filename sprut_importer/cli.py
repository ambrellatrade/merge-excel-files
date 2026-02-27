from __future__ import annotations

import argparse

from .core import run


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Build hierarchical SprutTP import file from Excel specs")
    parser.add_argument("--base", default=r"C:\CODEX\SPRUTTECHNOBASE")
    parser.add_argument("--head_dir", default=r"input\head")
    parser.add_argument("--specs_dir", default=r"input\specs")
    parser.add_argument("--template", default=r"template\sprut_template.xlsx")
    parser.add_argument("--out", default=r"output\sprut_import.xlsx")
    parser.add_argument("--report", default=r"output\report.xlsx")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return run(args)
