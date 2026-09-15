#!/usr/bin/env python3
"""
Export every Keap contact's tags to an Excel file, keyed by email,
so they can later be matched against and re-inserted into HubSpot.

Setup:
    pip install requests openpyxl python-dotenv
    Create a .env file next to this script containing:
        KEAP_API_KEY=the key you generated

Usage:
    python keap_export_tags.py [--output keap_tags_export.xlsx]

Auth: Keap REST API v1, key-based auth via the `X-Keap-API-Key` header.
Docs: https://developer.infusionsoft.com/docs/rest/
"""

import os
import sys
import time
import argparse
import requests
from openpyxl import Workbook

try:
    from dotenv import load_dotenv
    load_dotenv()  # reads .env in the current directory (or a parent) into os.environ
except ImportError:
    sys.exit("Missing dependency: run `pip install python-dotenv` (or `pip install -r requirements.txt`).")

BASE_URL = "https://api.infusionsoft.com/crm/rest/v1"
PAGE_SIZE = 200  # practical max page size Keap's /contacts endpoint accepts


def get_api_key():
    key = os.environ.get("KEAP_API_KEY")
    if not key:
        sys.exit("Set KEAP_API_KEY in a .env file (or as an env var) before running this.")
    return key


def keap_get(session, path, params=None, max_retries=5):
    """GET against the Keap API with basic 429/5xx retry + backoff."""
    url = f"{BASE_URL}{path}" if path.startswith("/") else path
    resp = None
    for attempt in range(max_retries):
        resp = session.get(url, params=params, timeout=30)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 5) or 5)
            print(f"  Rate limited, waiting {wait}s...")
            time.sleep(wait)
            continue
        if resp.status_code >= 500:
            wait = 2 ** attempt
            print(f"  Server error {resp.status_code}, retrying in {wait}s...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()


def fetch_all_tags(session):
    """Fetch every tag definition (id -> name) to label contacts' tag_ids."""
    tags = {}
    offset = 0
    while True:
        data = keap_get(session, "/tags", params={"limit": 1000, "offset": offset})
        batch = data.get("tags", [])
        for t in batch:
            tags[t["id"]] = t.get("name", "")
        offset += len(batch)
        if len(batch) < 1000:
            break
    print(f"Fetched {len(tags)} tag definitions.")
    return tags


def fetch_all_contacts_with_tags(session):
    """
    Fetch every contact along with its tag_ids in one pass, using
    optional_properties=tag_ids so we don't need a separate API call
    per contact just to get its tags.
    """
    contacts = []
    offset = 0
    while True:
        params = {
            "limit": PAGE_SIZE,
            "offset": offset,
            "optional_properties": "tag_ids",
        }
        data = keap_get(session, "/contacts", params=params)
        batch = data.get("contacts", [])
        contacts.extend(batch)
        print(f"  Fetched {len(contacts)} contacts so far...")
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return contacts


def primary_email(contact):
    emails = contact.get("email_addresses") or []
    for e in emails:
        if e.get("field") == "EMAIL1":
            return e.get("email", "")
    return emails[0]["email"] if emails else ""


def main():
    parser = argparse.ArgumentParser(description="Export Keap contacts + tags to Excel.")
    parser.add_argument("--output", default="keap_tags_export.xlsx")
    args = parser.parse_args()

    api_key = get_api_key()
    session = requests.Session()
    session.headers.update({"X-Keap-API-Key": api_key, "Accept": "application/json"})

    print("Fetching tag definitions...")
    tag_names = fetch_all_tags(session)

    print("Fetching contacts (this can take a while for large lists)...")
    contacts = fetch_all_contacts_with_tags(session)
    print(f"Fetched {len(contacts)} contacts total.")

    wb = Workbook()

    # Sheet 1: one row per contact, tags semicolon-joined.
    # Semicolon-separated is what HubSpot's import expects for a
    # multi-checkbox / multi-select property, if that's the route you
    # take when you reinsert these into HubSpot later.
    ws1 = wb.active
    ws1.title = "Contacts"
    ws1.append(["keap_contact_id", "email", "first_name", "last_name", "tags"])

    # Sheet 2: one row per contact-tag pair, for pivoting/spot-checking.
    ws2 = wb.create_sheet("Contact-Tag Pairs")
    ws2.append(["keap_contact_id", "email", "tag_id", "tag_name"])

    unmatched_tag_ids = set()
    skipped_no_email = 0

    for c in contacts:
        cid = c.get("id")
        email = primary_email(c)
        if not email:
            skipped_no_email += 1
        first = c.get("given_name", "") or ""
        last = c.get("family_name", "") or ""
        tag_ids = c.get("tag_ids") or []
        names = []
        for tid in tag_ids:
            name = tag_names.get(tid)
            if name is None:
                unmatched_tag_ids.add(tid)
                name = f"[unknown tag {tid}]"
            names.append(name)
            ws2.append([cid, email, tid, name])
        ws1.append([cid, email, first, last, "; ".join(sorted(names))])

    wb.save(args.output)
    print(f"\nSaved {len(contacts)} contacts to {args.output}")
    if skipped_no_email:
        print(f"Note: {skipped_no_email} contact(s) had no email address (matching to HubSpot by email won't work for these).")
    if unmatched_tag_ids:
        print(f"Warning: {len(unmatched_tag_ids)} tag id(s) had no matching name: {sorted(unmatched_tag_ids)}")


if __name__ == "__main__":
    main()