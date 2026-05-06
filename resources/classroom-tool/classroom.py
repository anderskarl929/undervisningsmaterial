"""CLI för Classroom-verktyget. Anonymiserar elev-data via per-kurs-pseudonymer.

Användning:
  python classroom.py list                                   # Lista mina aktiva kurser
  python classroom.py list --all                             # Inkludera arkiverade
  python classroom.py summary <course-id>                    # Anonymiserad metadata-sammanställning
  python classroom.py summary <id> -o file.md                # Spara till fil
  python classroom.py key <course-id>                        # HTML-nyckel (alias -> namn)
  python classroom.py read <course-id> <work-id> <Elev N>    # Läs en elevs inlämning
  python classroom.py dump <course-id> <work-id>             # Läs hela klassens inlämningar
  python classroom.py cache --clear | --purge                # Hantera Drive-text-cache

Auth hanteras av `gws` (Google Workspace CLI). Logga in med `gws auth login`.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def cmd_list(args: argparse.Namespace) -> int:
    import gws_client

    courses = gws_client.list_courses(active_only=not args.all)
    if not courses:
        print("Inga kurser hittades.")
        return 0
    width = max(len(c.get("id", "")) for c in courses)
    for c in courses:
        cid = c.get("id", "").ljust(width)
        name = c.get("name", "")
        section = c.get("section", "")
        suffix = f"  ({section})" if section else ""
        print(f"{cid}  {name}{suffix}")
    return 0


def cmd_summary(args: argparse.Namespace) -> int:
    from summary import build_summary

    md = build_summary(args.course_id)
    if args.out:
        Path(args.out).write_text(md, encoding="utf-8")
        print(f"Skrev {args.out}", file=sys.stderr)
    else:
        sys.stdout.write(md)
    return 0


def cmd_key(args: argparse.Namespace) -> int:
    from key import build_key

    out_path = build_key(args.course_id)
    print(f"Nyckel skriven: {out_path}", file=sys.stderr)
    print("Öppna i browser, skriv ut fysiskt, radera filen efteråt.", file=sys.stderr)
    return 0


def cmd_read(args: argparse.Namespace) -> int:
    from submissions import build_read

    md = build_read(
        args.course_id,
        args.coursework_id,
        args.alias,
        use_cache=not args.no_cache,
    )
    if args.out:
        Path(args.out).write_text(md, encoding="utf-8")
        print(f"Skrev {args.out}", file=sys.stderr)
    else:
        sys.stdout.write(md)
    return 0


def cmd_dump(args: argparse.Namespace) -> int:
    from submissions import build_dump

    md = build_dump(args.course_id, args.coursework_id, use_cache=not args.no_cache)
    if args.out:
        Path(args.out).write_text(md, encoding="utf-8")
        print(f"Skrev {args.out}", file=sys.stderr)
    else:
        sys.stdout.write(md)
    return 0


def cmd_cache(args: argparse.Namespace) -> int:
    import cache

    if args.clear:
        n = cache.clear_all()
        print(f"Raderade {n} cachade filer.", file=sys.stderr)
        return 0
    n = cache.purge_expired()
    print(f"Raderade {n} expiderade cachade filer.", file=sys.stderr)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="classroom", description=__doc__)
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="Lista mina kurser")
    p_list.add_argument("--all", action="store_true", help="Inkludera ej aktiva kurser")
    p_list.set_defaults(func=cmd_list)

    p_sum = sub.add_parser("summary", help="Anonymiserad sammanställning")
    p_sum.add_argument("course_id", help="Kurs-ID (från `list`)")
    p_sum.add_argument("-o", "--out", help="Skriv till fil istället för stdout")
    p_sum.set_defaults(func=cmd_summary)

    p_key = sub.add_parser("key", help="Skriv ut HTML-nyckel (alias -> riktiga namn)")
    p_key.add_argument("course_id", help="Kurs-ID")
    p_key.set_defaults(func=cmd_key)

    p_read = sub.add_parser("read", help="Läs en elevs inlämning (anonymiserad)")
    p_read.add_argument("course_id", help="Kurs-ID")
    p_read.add_argument("coursework_id", help="Uppgifts-ID (från `summary`-länk eller API)")
    p_read.add_argument("alias", help="Alias, t.ex. 'Elev 7' eller bara '7'")
    p_read.add_argument("-o", "--out", help="Skriv till fil istället för stdout")
    p_read.add_argument("--no-cache", action="store_true", help="Hoppa över Drive-text-cache")
    p_read.set_defaults(func=cmd_read)

    p_dump = sub.add_parser("dump", help="Läs alla inlämningar för en uppgift")
    p_dump.add_argument("course_id", help="Kurs-ID")
    p_dump.add_argument("coursework_id", help="Uppgifts-ID")
    p_dump.add_argument("-o", "--out", help="Skriv till fil istället för stdout")
    p_dump.add_argument("--no-cache", action="store_true", help="Hoppa över Drive-text-cache")
    p_dump.set_defaults(func=cmd_dump)

    p_cache = sub.add_parser("cache", help="Hantera Drive-text-cache (cache/)")
    grp = p_cache.add_mutually_exclusive_group()
    grp.add_argument("--clear", action="store_true", help="Radera all cache")
    grp.add_argument("--purge", action="store_true", help="Radera expirerad cache (default)")
    p_cache.set_defaults(func=cmd_cache)

    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
