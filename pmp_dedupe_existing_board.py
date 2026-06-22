#!/usr/bin/env python3
"""Mark duplicate PMP candidates by Content Hash without deleting files."""

from __future__ import annotations

import json
import os
import sys
from collections import defaultdict
from pathlib import Path

import gspread
from dotenv import load_dotenv

from pmp_resume_watcher import (
    CANDIDATE_LOOKUP_HEADERS,
    SCORING_BOARD_HEADERS,
    SELECTION_BOARD_HEADERS,
    Settings,
    get_google_credentials,
    rebuild_selection_board,
)


def load_settings() -> Settings:
    load_dotenv()
    script_dir = Path(__file__).resolve().parent
    return Settings(
        openai_api_key=os.getenv("OPENAI_API_KEY", ""),
        google_credentials_json=Path(os.getenv("GOOGLE_CREDENTIALS_JSON", str(script_dir / "credentials.json"))).expanduser().resolve(),
        google_token_pickle=Path(os.getenv("GOOGLE_TOKEN_PICKLE", str(script_dir / "token.pickle"))).expanduser().resolve(),
        google_sheet_id=os.environ["GOOGLE_SHEET_ID"],
        watch_folder=Path(os.environ["WATCH_FOLDER"]).expanduser().resolve(),
        worksheet_name=os.getenv("WORKSHEET_NAME", "Candidate Scoring Board"),
        drive_upload_folder_id=os.getenv("DRIVE_UPLOAD_FOLDER_ID") or None,
        google_auth_mode=os.getenv("GOOGLE_AUTH_MODE", "oauth").lower(),
    )


def dedupe_existing_board() -> dict:
    settings = load_settings()
    creds = get_google_credentials(settings)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(settings.google_sheet_id)
    lookup = spreadsheet.worksheet("Candidate Lookup")
    scoring = spreadsheet.worksheet("Candidate Scoring Board")
    selection = spreadsheet.worksheet("Selection Board")

    rows = lookup.get_all_values()
    header = rows[0] if rows else []
    if header[: len(CANDIDATE_LOOKUP_HEADERS)] != CANDIDATE_LOOKUP_HEADERS:
        lookup.update("A1:H1", [CANDIDATE_LOOKUP_HEADERS])

    groups: dict[str, list[tuple[int, list[str]]]] = defaultdict(list)
    for row_number, row in enumerate(rows[1:], start=2):
        content_hash = row[5].strip() if len(row) > 5 else ""
        if content_hash:
            groups[content_hash].append((row_number, row))

    marked = []
    scoring_ids = scoring.col_values(1)
    for content_hash, items in groups.items():
        if len(items) < 2:
            continue
        items.sort(key=lambda item: item[0])
        keeper_id = items[0][1][0]
        for row_number, row in items[1:]:
            duplicate_id = row[0]
            status = f"DUPLICATE_OF_{keeper_id}"
            lookup.update_cell(row_number, 8, status)
            for scoring_row, candidate_id in enumerate(scoring_ids, start=1):
                if candidate_id == duplicate_id:
                    scoring.update_cell(scoring_row, 30, status)
                    break
            marked.append({"duplicate_id": duplicate_id, "duplicate_of": keeper_id, "hash": content_hash})

    rebuild_selection_board(selection, scoring)
    return {"marked_duplicates": marked}


if __name__ == "__main__":
    print(json.dumps(dedupe_existing_board(), indent=2))
