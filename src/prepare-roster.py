from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from datetime import date, datetime, timedelta
import hashlib
from pathlib import Path


DEFAULT_INPUT = Path("config/roster.csv")
DEFAULT_DOB_START = date(2000, 1, 1)
DEFAULT_DOB_END = date(2004, 12, 31)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Add deterministic DOB values to a names-only roster CSV."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help="Input roster CSV, typically with Last,First columns.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output CSV path. Defaults to overwriting the input file.",
    )
    parser.add_argument(
        "--dob-start",
        default=DEFAULT_DOB_START.isoformat(),
        help="Inclusive lower bound for generated DOBs in YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--dob-end",
        default=DEFAULT_DOB_END.isoformat(),
        help="Inclusive upper bound for generated DOBs in YYYY-MM-DD format.",
    )
    return parser.parse_args()


def parse_iso_date(raw: str) -> date:
    return datetime.strptime(raw, "%Y-%m-%d").date()


def normalize_header(header: str) -> str:
    return "".join(ch for ch in header.strip().lower() if ch.isalnum())


def read_rows(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError(f"Roster file has no headers: {path}")
        headers = [header.strip() for header in reader.fieldnames if header]
        rows = [{k.strip(): (v or "").strip() for k, v in row.items() if k} for row in reader]
    return headers, rows


def row_name_fields(row: dict[str, str]) -> tuple[str, str]:
    normalized = {normalize_header(k): v for k, v in row.items()}
    last = normalized.get("last") or normalized.get("lastname")
    first = normalized.get("first") or normalized.get("firstname")
    if not first or not last:
        raise ValueError("Roster rows must include Last and First columns")
    return last, first


def row_dob_value(row: dict[str, str]) -> str:
    normalized = {normalize_header(k): v for k, v in row.items()}
    return normalized.get("dob") or normalized.get("dateofbirth") or ""


def deterministic_dob(last: str, first: str, occurrence: int, start: date, span_days: int) -> date:
    seed = f"{last.strip().casefold()}|{first.strip().casefold()}|{occurrence}".encode("utf-8")
    digest = hashlib.blake2b(seed, digest_size=8).digest()
    offset = int.from_bytes(digest, "big") % span_days
    return start + timedelta(days=offset)


def generate_dobs(rows: list[dict[str, str]], start: date, end: date) -> list[dict[str, str]]:
    if end < start:
        raise ValueError("dob-end must be on or after dob-start")

    span_days = (end - start).days + 1
    seen_tuples: set[tuple[str, str, str]] = set()
    name_counts: defaultdict[tuple[str, str], int] = defaultdict(int)
    output_rows: list[dict[str, str]] = []

    for row in rows:
        last, first = row_name_fields(row)
        normalized_name = (last.strip().casefold(), first.strip().casefold())
        existing_dob = row_dob_value(row)

        if existing_dob:
            dob = parse_iso_date(existing_dob).isoformat()
        else:
            occurrence = name_counts[normalized_name]
            name_counts[normalized_name] += 1
            generated = deterministic_dob(last=last, first=first, occurrence=occurrence, start=start, span_days=span_days)
            dob = generated.isoformat()
            while (normalized_name[0], normalized_name[1], dob) in seen_tuples:
                generated = generated + timedelta(days=1)
                if generated > end:
                    generated = start
                dob = generated.isoformat()

        seen_tuples.add((normalized_name[0], normalized_name[1], dob))
        output_rows.append({"Last": last, "First": first, "DOB": dob})

    return output_rows


def write_rows(path: Path, rows: list[dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=["Last", "First", "DOB"])
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve() if args.output else input_path
    start = parse_iso_date(args.dob_start)
    end = parse_iso_date(args.dob_end)

    _, rows = read_rows(input_path)
    normalized_rows = generate_dobs(rows=rows, start=start, end=end)
    write_rows(output_path, normalized_rows)
    print(f"Roster normalized with DOBs: {output_path}")


if __name__ == "__main__":
    main()
