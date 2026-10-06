import argparse
import os
import random
import re
import shutil
import sys
import time
from concurrent.futures import ProcessPoolExecutor
from datetime import datetime, timedelta
from itertools import repeat
from pathlib import Path

from openpyxl import Workbook

EXCEL_MAX_DATA_ROWS = 1_048_575
FREE_SPACE_RESERVE = 10**9
SIZE_UNITS = {"B": 1, "KB": 10**3, "MB": 10**6, "GB": 10**9, "TB": 10**12}
HEADER = [
    "order_id",
    "customer_name",
    "region",
    "product",
    "quantity",
    "unit_price",
    "order_date",
]
REGIONS = ["North", "South", "East", "West", "Central"]
PRODUCTS = ["Laptop", "Monitor", "Keyboard", "Mouse", "Headset", "Webcam", "Dock"]


def parse_size(text: str) -> int:
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*([KMGT]?B)", text.strip().upper())
    if not match:
        raise argparse.ArgumentTypeError(
            f"invalid size {text!r}, use for example 500MB, 20GB or 2TB"
        )
    return int(float(match.group(1)) * SIZE_UNITS[match.group(2)])


def positive_int(text: str) -> int:
    value = int(text)
    if value < 1:
        raise argparse.ArgumentTypeError("must be 1 or more")
    return value


def sheet_rows(text: str) -> int:
    value = int(text)
    if not 1 <= value <= EXCEL_MAX_DATA_ROWS:
        raise argparse.ArgumentTypeError(
            f"must be between 1 and {EXCEL_MAX_DATA_ROWS:,} "
            "(Excel's row limit minus the header row)"
        )
    return value


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create .xlsx workbooks with several sheets of order data."
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        required=True,
        help="directory for the workbooks; must not already contain a dataset",
    )
    parser.add_argument(
        "--files",
        type=positive_int,
        default=2,
        help="unique workbooks to generate (default: %(default)s)",
    )
    parser.add_argument(
        "--sheets-per-file",
        type=positive_int,
        default=4,
        help="worksheets per workbook (default: %(default)s)",
    )
    parser.add_argument(
        "--rows-per-sheet",
        type=sheet_rows,
        default=100_000,
        help="data rows per worksheet, at most 1,048,575 (default: %(default)s)",
    )
    parser.add_argument(
        "--target-size",
        type=parse_size,
        help="copy the generated workbooks until the dataset reaches this size, "
        "for example 50GB or 2TB",
    )
    parser.add_argument(
        "--workers",
        type=positive_int,
        default=min(4, os.cpu_count() or 1),
        help="processes used to generate the unique workbooks (default: %(default)s)",
    )
    return parser.parse_args(argv)


def workbook_name(number: int) -> str:
    return f"dataset_{number:05d}.xlsx"


def write_workbook(path: Path, sheets: int, rows: int, first_id: int, seed: int) -> int:
    rng = random.Random(seed)
    first_day = datetime(2024, 1, 1)
    workbook = Workbook(write_only=True)
    order_id = first_id
    for sheet_number in range(1, sheets + 1):
        sheet = workbook.create_sheet(f"Sheet{sheet_number}")
        sheet.append(HEADER)
        for _ in range(rows):
            sheet.append(
                [
                    order_id,
                    f"Customer {rng.randrange(1_000_000):06d}",
                    rng.choice(REGIONS),
                    rng.choice(PRODUCTS),
                    rng.randint(1, 20),
                    round(rng.uniform(5, 2500), 2),
                    first_day + timedelta(seconds=rng.randrange(730 * 86400)),
                ]
            )
            order_id += 1
    workbook.save(path)
    return sheets * rows


def replicate(sources: list[Path], output_dir: Path, target_bytes: int, first_number):
    total = sum(path.stat().st_size for path in sources)
    copies = []
    while total < target_bytes:
        source = sources[len(copies) % len(sources)]
        target = output_dir / workbook_name(first_number + len(copies))
        shutil.copyfile(source, target)
        total += source.stat().st_size
        copies.append(target)
    return copies


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    output_dir = args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)
    if any(output_dir.glob("dataset_*.xlsx")):
        print(f"{output_dir} already contains dataset_*.xlsx files", file=sys.stderr)
        return 1

    started = time.perf_counter()
    paths = [output_dir / workbook_name(number) for number in range(1, args.files + 1)]
    rows_per_file = args.sheets_per_file * args.rows_per_sheet
    first_ids = [index * rows_per_file + 1 for index in range(args.files)]
    jobs = (
        paths,
        repeat(args.sheets_per_file),
        repeat(args.rows_per_sheet),
        first_ids,
        range(1, args.files + 1),
    )
    workers = min(args.workers, args.files)
    if workers == 1:
        list(map(write_workbook, *jobs))
    else:
        with ProcessPoolExecutor(max_workers=workers) as pool:
            list(pool.map(write_workbook, *jobs))
    elapsed = time.perf_counter() - started
    print(f"Generated {args.files} unique workbooks in {elapsed:.1f}s")

    copies = []
    if args.target_size:
        needed = args.target_size - sum(path.stat().st_size for path in paths)
        free = shutil.disk_usage(output_dir).free
        if needed > free - FREE_SPACE_RESERVE:
            print(
                f"Not enough disk space: {needed:,} more bytes are needed, "
                f"{free:,} bytes are free (1 GB is kept in reserve)",
                file=sys.stderr,
            )
            return 1
        copies = replicate(paths, output_dir, args.target_size, args.files + 1)
        print(f"Added {len(copies)} copies to reach the target size")

    files = paths + copies
    total_bytes = sum(path.stat().st_size for path in files)
    print(
        f"Dataset: {len(files):,} files, {len(files) * args.sheets_per_file:,} sheets, "
        f"{len(files) * rows_per_file:,} rows, {total_bytes:,} bytes "
        f"({total_bytes / 10**9:,.2f} GB) in {output_dir}"
    )
    if copies:
        print("Copied workbooks repeat the order_id values of the workbook they copy.")
    print(f"Finished in {time.perf_counter() - started:.1f}s")
    return 0


if __name__ == "__main__":
    sys.exit(main())
