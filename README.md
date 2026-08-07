# XRF Uploader

Parses XRF instrument `.txt` exports, generates per-sample Excel reports, and uploads results to Benchling.

## Setup

```bash
pip install -r requirements.txt
```

Create a `dat/` directory and add your Benchling API key:

```bash
mkdir -p dat
```

`dat/secrets.json`:
```json
{"AUTHOR_API_KEY": "your_api_key_here"}
```

`dat/` is gitignored — this file will never be committed.

Alternatively, set the environment variable `BENCHLING_API_KEY` instead.

## Usage

```bash
# Preview parsing and calibration without uploading
python run_xrf.py path/to/file.txt --dry-run

# Full upload — generates xlsx reports and uploads to Benchling
python run_xrf.py path/to/file.txt

# Write xlsx reports to a specific directory
python run_xrf.py path/to/file.txt --output-dir /path/to/reports/
```

## What it does

1. **Parses** the `.txt` file — identifies YP50F-0..6 calibration standards (first occurrence) and unknown samples.
2. **Calibrates** — fits two-range linear curves per element (18 calibrated elements: K, Ca, V, Cr, Mn, Fe, Co, Ni, Cu, Zn, Ga, Rb, Sr, Mo, Cd, Tl, Pb, Bi).
3. **Assigns signal quality** — `No Signal`, `Not Significant`, `Low Signal`, `In Range`, or `High Signal`.
4. **Generates xlsx reports** — one per sample, with `Samdat`, `Caldat`, and per-element sheets containing Benchling Copyable Tables.
5. **Uploads** to three Benchling assay schemas:
   - `assaysch_9LFkZwZL` — element concentrations
   - `assaysch_e7BnMa7Z` — element signal intensities
   - `assaysch_dHcQUAwe` — run summary

## Configuration

Benchling field name slugs are defined as constants at the top of `xrf_upload.py` (`FIELD_CONC_*`, `FIELD_SIG_*`, `FIELD_SUM_*`). Verify these against your actual Benchling schema fields before the first live upload.
