#!/usr/bin/env python3
"""
Export every Keap company's tags to an Excel file.

Keap has no native "tag a company" feature -- tags only ever attach to
Contacts (see https://community.keap.com/t/how-to-add-a-tag-to-a-company/34877).
So a company's "tags" here are the UNION of the tags carried by every
contact linked to that company (via each contact's `company` field).

Setup:
    pip install requests openpyxl python-dotenv
    Uses the same .env as "keap tags.py":
        KEAP_API_KEY=the key you generated

Usage:
    python keap_company_tags.py [--output keap_company_tags_export.xlsx]

Auth: Keap REST API v1, key-based auth via the `X-Keap-API-Key` header.
Docs: https://developer.infusionsoft.com/docs/rest/

NOTE on unverified assumptions (same caveat style as "get subs.py" --
these match the v1 docs but haven't been re-confirmed against a live
response as of this writing, so double check if something 404s or a
field comes back empty):
    - GET /companies exists in v1, paginates with limit/offset, and the
      list key in the response is "companies". If your Keap account's
      REST API doesn't expose /companies yet, delete the
      fetch_all_companies() call in main() and the script will fall
      back to only companies discovered through tagged contacts.
    - A contact's linked company comes back as a nested object under
      the "company" key, e.g. {"id": 123, "company_name": "Acme"}. If
      that field is null/missing by default, pass
      optional_properties=tag_ids,company (already done below).
"""

import os
import re
import html
import sys
import time
import argparse
import requests
from openpyxl import Workbook

try:
    from dotenv import load_dotenv
    load_dotenv()
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
    """Fetch every tag definition (id -> name)."""
    tags = {}
    offset = 0
    while True:
        data = keap_get(session, "/tags", params={"limit": 1000, "offset": offset})
        batch = data.get("tags", [])
        for t in batch:
            tags[t["id"]] = html.unescape(t.get("name", "") or "")
        offset += len(batch)
        if len(batch) < 1000:
            break
    print(f"Fetched {len(tags)} tag definitions.")
    return tags


def slugify_tag(name):
    """Best-effort HubSpot internal-name suggestion (lowercase, underscores).
    HubSpot auto-generates its own internal value when you paste labels in,
    so treat this as a preview/sanity-check rather than the source of truth."""
    slug = re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")
    return slug or "tag"


def write_tag_reference(tag_names, wb, txt_path):
    """
    Write every distinct Keap tag name to:
      1. A plain .txt file, one tag per line, sorted -- paste this directly
         into HubSpot's "Add options" box when creating a multi-checkbox or
         dropdown select property to pre-populate every option at once.
      2. An "All Tags" sheet in the same workbook, with a suggested
         HubSpot internal value alongside each label, for reference.
    Returns the sorted list of unique tag names.
    """
    unique_names = sorted({n for n in tag_names.values() if n}, key=str.lower)

    with open(txt_path, "w", encoding="utf-8") as f:
        for name in unique_names:
            f.write(name + "\n")
    print(f"Wrote {len(unique_names)} unique tag names to {txt_path}")

    ws = wb.create_sheet("All Tags")
    ws.append(["tag_name (HubSpot label)", "suggested_internal_value"])
    seen_slugs = {}
    for name in unique_names:
        slug = slugify_tag(name)
        if slug in seen_slugs:
            seen_slugs[slug] += 1
            slug = f"{slug}_{seen_slugs[slug]}"
        else:
            seen_slugs[slug] = 0
        ws.append([name, slug])

    return unique_names


def fetch_all_contacts_with_tags_and_company(session):
    """Fetch every contact along with its tag_ids and linked company."""
    contacts = []
    offset = 0
    while True:
        params = {
            "limit": PAGE_SIZE,
            "offset": offset,
            "optional_properties": "tag_ids,company",
        }
        data = keap_get(session, "/contacts", params=params)
        batch = data.get("contacts", [])
        contacts.extend(batch)
        print(f"  Fetched {len(contacts)} contacts so far...")
        if len(batch) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
    return contacts


def fetch_all_companies(session):
    """Fetch every company (id -> name), so companies with zero tagged
    contacts still show up in the export. Returns {} if the endpoint
    isn't available on this account (caught and reported in main())."""
    companies = {}
    offset = 0
    while True:
        data = keap_get(session, "/companies", params={"limit": PAGE_SIZE, "offset": offset})
        batch = data.get("companies", [])
        for c in batch:
            companies[c["id"]] = c.get("company_name", "")
        offset += len(batch)
        if len(batch) < PAGE_SIZE:
            break
    print(f"Fetched {len(companies)} companies.")
    return companies


def main():
    parser = argparse.ArgumentParser(description="Export Keap companies + their contacts' tags to Excel.")
    parser.add_argument("--output", default="keap_company_tags_export.xlsx")
    parser.add_argument("--tags-txt", default="keap_all_tags_for_hubspot.txt",
                         help="Plain-text file listing every distinct tag name, one per line.")
    args = parser.parse_args()

    api_key = get_api_key()
    session = requests.Session()
    session.headers.update({"X-Keap-API-Key": api_key, "Accept": "application/json"})

    print("Fetching tag definitions...")
    tag_names = fetch_all_tags(session)

    print("Fetching companies...")
    try:
        company_names = fetch_all_companies(session)
    except requests.HTTPError as e:
        print(f"  Warning: /companies lookup failed ({e}); falling back to "
              f"company names seen on contacts only.")
        company_names = {}

    print("Fetching contacts (tags + linked company)...")
    contacts = fetch_all_contacts_with_tags_and_company(session)
    print(f"Fetched {len(contacts)} contacts total.")

    # company_id -> {"name": str, "tag_ids": set(), "contact_count": int}
    companies = {}
    contacts_no_company = 0

    for c in contacts:
        company = c.get("company") or {}
        company_id = company.get("id")
        if not company_id:
            contacts_no_company += 1
            continue
        entry = companies.setdefault(company_id, {
            "name": company.get("company_name") or company_names.get(company_id, ""),
            "tag_ids": set(),
            "contact_count": 0,
        })
        entry["contact_count"] += 1
        for tid in (c.get("tag_ids") or []):
            entry["tag_ids"].add(tid)

    # Include companies that exist but had no linked/tagged contacts at all.
    for cid, name in company_names.items():
        companies.setdefault(cid, {"name": name, "tag_ids": set(), "contact_count": 0})

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Companies"
    ws1.append(["keap_company_id", "company_name", "contact_count", "tags"])

    write_tag_reference(tag_names, wb, args.tags_txt)

    ws2 = wb.create_sheet("Company-Tag Pairs")
    ws2.append(["keap_company_id", "company_name", "tag_id", "tag_name"])

    unmatched_tag_ids = set()

    for cid, info in companies.items():
        names = []
        for tid in info["tag_ids"]:
            name = tag_names.get(tid)
            if name is None:
                unmatched_tag_ids.add(tid)
                name = f"[unknown tag {tid}]"
            names.append(name)
            ws2.append([cid, info["name"], tid, name])
        ws1.append([cid, info["name"], info["contact_count"], "; ".join(sorted(names))])

    wb.save(args.output)
    print(f"\nSaved {len(companies)} companies to {args.output}")
    if contacts_no_company:
        print(f"Note: {contacts_no_company} contact(s) had no linked company.")
    if unmatched_tag_ids:
        print(f"Warning: {len(unmatched_tag_ids)} tag id(s) had no matching name: {sorted(unmatched_tag_ids)}")


if __name__ == "__main__":
    main()
