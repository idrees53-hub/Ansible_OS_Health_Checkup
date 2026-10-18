#!/usr/bin/env python3

from pathlib import Path
import subprocess
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent

DATA_DIR = PROJECT_ROOT / "reports" / "2026-Q4" / "data"
OUTPUT_DIR = PROJECT_ROOT / "reports" / "2026-Q4" / "servers"

SINGLE_REPORT_GENERATOR = (
    PROJECT_ROOT
    / "report_generator"
    / "generate_server_report.py"
)


def main():

    print("=" * 60)
    print("STC RHEL OS HEALTH CHECK")
    print("Batch DOCX Report Generator")
    print("=" * 60)

    if not DATA_DIR.exists():

        print(f"ERROR: Data directory does not exist:")
        print(DATA_DIR)

        sys.exit(1)

    if not SINGLE_REPORT_GENERATOR.exists():

        print("ERROR: Individual report generator not found:")
        print(SINGLE_REPORT_GENERATOR)

        sys.exit(1)

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    json_files = sorted(
        DATA_DIR.glob("*.json")
    )

    if not json_files:

        print("No JSON report files found.")
        print(DATA_DIR)

        sys.exit(1)

    print(f"JSON files found: {len(json_files)}")
    print()

    successful = 0
    failed = 0

    for json_file in json_files:

        output_file = (
            OUTPUT_DIR
            / f"{json_file.stem}.docx"
        )

        print(
            f"Generating: {json_file.name}"
        )

        result = subprocess.run(
            [
                sys.executable,
                str(SINGLE_REPORT_GENERATOR),
                str(json_file),
                str(output_file),
            ],
            capture_output=True,
            text=True,
        )

        if result.returncode == 0:

            print(
                f"  SUCCESS -> {output_file.name}"
            )

            successful += 1

        else:

            print(
                f"  FAILED -> {json_file.name}"
            )

            if result.stderr:
                print(result.stderr)

            failed += 1

    print()
    print("=" * 60)
    print("BATCH GENERATION SUMMARY")
    print("=" * 60)

    print(f"JSON files found : {len(json_files)}")
    print(f"Reports created  : {successful}")
    print(f"Reports failed   : {failed}")

    print()
    print(f"Output directory:")
    print(OUTPUT_DIR)

    if failed > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
