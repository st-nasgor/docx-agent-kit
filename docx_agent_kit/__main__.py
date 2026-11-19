"""Command line: python -m docx_agent_kit {inspect,manifest,render} ..."""

import argparse
import json
import sys


def _print(value):
    print(json.dumps(value, indent=2, ensure_ascii=False, default=str))


def main(argv=None):
    from . import inspect, manifest_from_template, render

    parser = argparse.ArgumentParser(prog="docx-agent-kit", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    p_inspect = sub.add_parser("inspect", help="report features used by a docx")
    p_inspect.add_argument("docx")

    p_manifest = sub.add_parser("manifest", help="list template fields")
    p_manifest.add_argument("docx")

    p_render = sub.add_parser("render", help="fill a template from JSON")
    p_render.add_argument("docx")
    p_render.add_argument("payload", help="JSON file or '-' for stdin")
    p_render.add_argument("-o", "--out", default=None)

    args = parser.parse_args(argv)

    if args.command == "inspect":
        _print(inspect(args.docx))
    elif args.command == "manifest":
        _print(manifest_from_template(args.docx))
    elif args.command == "render":
        if args.payload == "-":
            payload = json.load(sys.stdin)
        else:
            with open(args.payload, encoding="utf-8") as handle:
                payload = json.load(handle)
        _print({"out": render(args.docx, payload, args.out)})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
