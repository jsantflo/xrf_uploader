"""
Entry point for the XRF uploader.

Usage:
    python run_xrf.py path/to/file.txt [--dry-run] [--output-dir DIR]

The Benchling API key is read from a secrets.json file in the same directory:
    {"AUTHOR_API_KEY": "your_key_here"}

Or set the BENCHLING_API_KEY environment variable instead.

Outputs:
  - One .xlsx report per sample, written to --output-dir (default: a folder
    named after the .txt file, placed alongside it).
  - Assay results uploaded to Benchling for each sample.
"""

import argparse
import json
import os
import sys
from pathlib import Path

from benchling_sdk.benchling import Benchling
from benchling_sdk.auth.api_key_auth import ApiKeyAuth

from xrf_upload import (
    upload_xrf_from_txt, parse_xrf_txt,
    find_entity_by_name, find_xrf_entity_by_name,
    _parse_replicate_suffix,
)

BENCHLING_URL = "https://florrent.benchling.com"
SECRETS_FILE  = Path(__file__).parent / "dat" / "secrets.json"


def load_api_key() -> str:
    if SECRETS_FILE.exists():
        with open(SECRETS_FILE) as f:
            return json.load(f)["AUTHOR_API_KEY"]
    if "BENCHLING_API_KEY" in os.environ:
        return os.environ["BENCHLING_API_KEY"]
    print(
        "Error: no API key found. "
        f"Expected {SECRETS_FILE} with {{\"AUTHOR_API_KEY\": \"...\"}} "
        "or set the BENCHLING_API_KEY environment variable.\n"
    )
    sys.exit(1)


def resolve_entity_interactive(benchling, sample_name: str, dry_run: bool = False):
    """
    Resolve a Benchling custom entity for sample_name.
    dry_run=True: search only, no interactive prompt — returns entity ID or None.
    dry_run=False: loops interactively until a match is found or the user skips.
    Returns entity ID string, or None if not found / skipped.
    """
    base_name, replicate_n = _parse_replicate_suffix(sample_name)
    if replicate_n is not None:
        print(f"    Inter-sample replicate #{replicate_n} of '{base_name}'")

    name_to_check = base_name
    while True:
        entity = find_entity_by_name(benchling, name_to_check)
        if entity is not None:
            url = getattr(entity, "web_url", None) or f"{BENCHLING_URL}/custom-entities/{entity.id}"
            print(f"    Matched: {entity.name} ({entity.id})  →  {url}")
            return entity.id

        if dry_run:
            print(f"    [DRY RUN] No entity found for '{name_to_check}' — will block live upload.")
            return None

        print(f"\n  Warning: no Benchling entity found matching '{name_to_check}'.")
        while True:
            choice = input("  [1] Try a different entity name  [2] Skip this sample: ").strip()
            if choice in ("1", "2"):
                break
            print("  Please enter 1 or 2.")

        if choice == "2":
            print(f"  Skipping entity link for '{sample_name}'.")
            return None

        name_to_check = input("  Entity name to search: ").strip()


def resolve_xrf_entity_interactive(benchling, sample_name: str, replicate_n: int,
                                    dry_run: bool = False):
    """
    Before creating a new XRF Sample entity for an inter-sample replicate, check whether
    one with the same name already exists.  If it does, ask the user whether this run is:
      [1] The same physical sample (reuse the existing entity — new results are added to it)
      [2] A new sample with an accidentally repeated sampling number (create a new entity;
          the user should correct the name in the input file before re-running)

    Returns the existing entity ID to reuse, or None to signal that a new entity should be
    created by upload_xrf_from_txt.
    """
    existing = find_xrf_entity_by_name(benchling, sample_name)
    if existing is None:
        return None  # no conflict — create fresh

    url = getattr(existing, "web_url", None) or f"{BENCHLING_URL}/custom-entities/{existing.id}"
    print(f"\n  Warning: an XRF Sample entity named '{sample_name}' already exists.")
    print(f"  Existing: {existing.name} ({existing.id})  →  {url}")

    if dry_run:
        print(f"  [DRY RUN] Would ask: same physical sample or repeated sampling number?")
        print(f"  [DRY RUN] Assuming reuse for preview.")
        return existing.id

    print(f"  Is this run:")
    print(f"    [1] The same physical sample — add these measurements to the existing entity")
    print(f"    [2] A new sample with an accidentally repeated sampling number")
    while True:
        choice = input("  Select (1 or 2): ").strip()
        if choice in ("1", "2"):
            break
        print("  Please enter 1 or 2.")

    if choice == "1":
        print(f"  Reusing existing entity: {existing.id}")
        return existing.id

    # choice == "2": new sample, wrong number — ask for a correction
    base_name, _ = _parse_replicate_suffix(sample_name)
    print(f"  Please enter the correct sampling number for this new sample")
    print(f"  (or press Enter to create a new entity anyway with sampling #{replicate_n}):")
    raw = input(f"  Correct sampling number [{replicate_n}]: ").strip()
    if raw.isdigit():
        corrected_n = int(raw)
        corrected_name = f"{base_name} {corrected_n}"
        print(f"  Will create new entity named '{corrected_name}' — "
              f"please update the input file accordingly.")
        # Return a sentinel tuple so main() can override the name
        return ("NEW", corrected_name, corrected_n)
    else:
        print(f"  Creating new entity with the original name '{sample_name}'.")
        return None  # create new entity with original name


def main():
    parser = argparse.ArgumentParser(
        description="Parse an XRF .txt file, generate xlsx reports, and upload results to Benchling."
    )
    parser.add_argument("txt", help="Path to the XRF instrument .txt export file.")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Parse and preview calibration/concentrations without writing to Benchling.",
    )
    parser.add_argument(
        "--output-dir", default=None,
        help="Directory to write xlsx reports into. "
             "Defaults to a folder named after the .txt file alongside it.",
    )
    args = parser.parse_args()

    txt_path = Path(args.txt)
    if not txt_path.exists():
        print(f"Error: file not found: {txt_path}")
        sys.exit(1)

    output_dir = Path(args.output_dir) if args.output_dir else txt_path.parent / txt_path.stem

    benchling = Benchling(url=BENCHLING_URL, auth_method=ApiKeyAuth(load_api_key()))

    entity_map = {}
    xrf_entity_map = {}   # sample_name -> existing XRF entity ID to reuse (replicates only)
    xrf_name_overrides = {}  # sample_name -> corrected name when user fixes a bad repeat number
    parsed_preview = parse_xrf_txt(txt_path)
    sample_names = list(parsed_preview["samples"].keys())
    if sample_names:
        label = "[DRY RUN] " if args.dry_run else ""
        print(f"\n{label}Resolving Benchling entities for {len(sample_names)} sample(s)...")
        for name in sample_names:
            print(f"  Searching for '{name}'...")
            entity_map[name] = resolve_entity_interactive(benchling, name, dry_run=args.dry_run)

            # For inter-sample replicates, check for an existing XRF entity with the same name
            _, replicate_n = _parse_replicate_suffix(name)
            if replicate_n is not None:
                result = resolve_xrf_entity_interactive(
                    benchling, name, replicate_n, dry_run=args.dry_run
                )
                if isinstance(result, tuple) and result[0] == "NEW":
                    # User corrected the sampling number — record the new name
                    _, corrected_name, _ = result
                    xrf_name_overrides[name] = corrected_name
                elif result is not None:
                    xrf_entity_map[name] = result

    results = upload_xrf_from_txt(
        txt_path=txt_path,
        output_dir=output_dir,
        benchling=benchling,
        dry_run=args.dry_run,
        entity_map=entity_map,
        xrf_entity_map=xrf_entity_map,
        xrf_name_overrides=xrf_name_overrides,
    )

    if not args.dry_run:
        uploaded   = [r for r in results if not r.get("skipped") and not r.get("duplicate")]
        skipped    = [r for r in results if r.get("skipped")]
        duplicates = [r for r in results if r.get("duplicate")]
        print(f"\nUpload complete — {len(uploaded)} uploaded, {len(skipped)} skipped, {len(duplicates)} duplicate(s) blocked.")
        for r in uploaded:
            print(f"\n  Sample         : {r['sample_name']}")
            print(f"  Summary result : {r['summary_result_id']}")
            print(f"  Conc. results  : {len(r['concentration_result_ids'])}")
            print(f"  Signal results : {len(r['signal_result_ids'])}")
        for r in skipped:
            print(f"\n  Skipped        : {r['sample_name']}")
        for r in duplicates:
            print(f"\n  Duplicate      : {r['sample_name']}")
    else:
        print(f"\n[DRY RUN] Reports written to: {output_dir}")
        print("No data was uploaded to Benchling.")


if __name__ == "__main__":
    main()
