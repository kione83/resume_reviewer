# PMP Resume Watcher

This script watches a PMP intake folder, batches applicant resumes, randomizes the processing order, extracts PDF/DOCX text, scores each resume with OpenAI, appends raw scores and rationale to Google Sheets, and moves each file to `processed/`, `duplicates/`, or `failed/`.

## Setup

Install dependencies:

```bash
pip install -r requirements.txt
```

Create a real `.env` from `.env.example`:

```bash
cp .env.example .env
```

Fill in:

- `OPENAI_API_KEY`: your OpenAI API key.
- `GOOGLE_SHEET_ID`: the target Google Sheet ID.
- `WATCH_FOLDER`: batch root folder. Resumes can be downloaded into `WATCH_FOLDER/intake`; top-level PDF/DOCX drops in `WATCH_FOLDER` are automatically moved into `intake`.
- `WORKSHEET_NAME`: defaults to `Candidate Scoring Board`.
- `DRIVE_UPLOAD_FOLDER_ID`: optional Google Drive folder ID where accepted resume copies should be uploaded.
- `BATCH_IDLE_MINUTES`: optional idle window before automatic batch processing starts; defaults to `10`.

By default, Google auth uses the existing OAuth files in this folder:

- `credentials.json`
- `token.pickle`

For OAuth, keep `GOOGLE_AUTH_MODE=oauth`. To use a service account instead, set `GOOGLE_AUTH_MODE=service_account` and point `GOOGLE_CREDENTIALS_JSON` to the service account JSON file.

If `DRIVE_UPLOAD_FOLDER_ID` is set, resumes that pass resume detection are copied to that Drive folder after they have been assigned randomized Candidate IDs.

## Folder Layout

The script creates and uses these folders under `WATCH_FOLDER`:

- `intake`: Steve downloads all resumes here.
- `staging`: batch processing workspace.
- `processed`: successfully processed resumes, renamed to Candidate IDs.
- `duplicates`: duplicate resumes detected by content hash.
- `failed`: files that fail processing.
- `ignored_non_resumes`: supported files that are not resumes.
- `batch_manifests`: JSON audit manifests for randomized batches.

## Run

```bash
python3 pmp_resume_watcher.py
```

Watch mode detects files in `WATCH_FOLDER/intake` and records intake metadata, but it does not immediately score them. If supported PDF/DOCX files are dropped directly into `WATCH_FOLDER`, watch mode moves them into `intake` first. It processes a randomized batch only after no new eligible files have arrived for `BATCH_IDLE_MINUTES`.

To process the current intake folder manually:

```bash
python3 pmp_resume_watcher.py --process-batch
```

The watcher ignores `.DS_Store`, hidden files, `~$` Office temp files, partial downloads, and unsupported file types. It keeps a `processed_files.json` hash log to avoid duplicate processing.

## Batch Randomization

Candidate IDs are assigned only during batch processing. Files are moved from `intake` to `staging`, shuffled with a recorded randomization seed, then renamed to Candidate IDs before scoring. This breaks the trail from email order to download order to Candidate ID order.

Each batch manifest records:

- Batch ID
- Randomization seed
- Original filename
- Intake timestamp
- Assigned Candidate ID
- Per-file processing timestamp
- Final status

## Resume Detection

Because the watch folder may receive non-resume documents, the script extracts text first and classifies the document before scoring it.

Default behavior is `RESUME_DETECTION_MODE=hybrid`:

- Obvious resumes are accepted by filename/content heuristics.
- Obvious non-resumes are moved to `ignored_non_resumes/`.
- Ambiguous documents are sent to OpenAI for a strict JSON resume/non-resume classification before scoring.

Optional modes:

- `heuristic`: never calls OpenAI for resume detection; only scores documents that pass local heuristics.
- `llm`: always asks OpenAI to classify supported PDF/DOCX files before scoring.

Ignored non-resumes are logged in `ignored_non_resumes/ignored_log.jsonl`. Scoring failures still go to `failed/` with `failed/error_log.jsonl`.

## Sheet Behavior

The v3 script writes:

- Candidate ID
- Candidate/name/email/file lookup with Content Hash and intake timestamp
- Three independent AI scoring passes
- Average score, spread, confidence, rationale, and selection status

Selection is rank/seat based rather than threshold based.

## Duplicate Handling

The watcher hashes normalized resume text during randomized batch processing. If the same content hash already exists in `processed_files.json` or `Candidate Lookup`, the file is moved to `duplicates/`, recorded in the batch manifest as a duplicate, and no new scoring-board row is created.

Controls:

- `ALLOW_REPROCESS=false`: default, skip duplicates.
- `ALLOW_REPROCESS=true`: update the existing Candidate ID row.
- `FORCE_NEW_CANDIDATE_ID=true`: with `ALLOW_REPROCESS=true`, create a new Candidate ID anyway.

Cleanup utility:

```bash
python3 pmp_dedupe_existing_board.py
```

This runs `dedupe_existing_board()`, groups `Candidate Lookup` rows by `Content Hash`, marks later duplicates as `DUPLICATE_OF_<Candidate ID>`, and rebuilds the `Selection Board`. It does not delete resume files.
