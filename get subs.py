

"""
Keap REST API v2 - export ALL subscriptions (any status) with joined contact
and product info, as a cross-check against the Recurring Order Search UI
report.

Field names below are confirmed from a real API response (paste sent back
2026-09-14) - no more guessing on the subscription object itself:
    id, quantity, active (bool), contact_id, product_id,
    subscription_plan_id, billing_amount, auto_charge, billing_frequency,
    billing_cycle ("DAY"/"WEEK"/"MONTH"/"YEAR"), start_date, last_bill_date,
    next_bill_date, merchant_account_id, payment_method_id, promo_code,
    reason_stopped, modification_time

Still unverified (not yet confirmed against a real response - if these
error out, check https://developer.infusionsoft.com/docs/restv2/ for the
actual shape):
    - the /subscriptions list endpoint's pagination shape
    - the /contacts/{id} response field names
    - the /products/{id} response field names

pip install requests
"""

import csv
import time
import requests

ACCESS_TOKEN = "KeapAK-6b66a3b3a6724a5acc3d66ee4a95f1b502d8733ec8847d63e6"
BASE_URL = "https://api.infusionsoft.com/crm/rest/v2"  # VERIFY this matches your dev portal
HEADERS = {"Authorization": f"Bearer {ACCESS_TOKEN}"}

OUTPUT_FILE = "keap_all_subscriptions.csv"


def get_all_subscriptions():
    """Page through ALL subscriptions, any status."""
    subscriptions = []
    url = f"{BASE_URL}/subscriptions"
    params = {"limit": 200}  # VERIFY max page size allowed

    while url:
        resp = requests.get(url, headers=HEADERS, params=params)
        resp.raise_for_status()
        data = resp.json()

        # VERIFY: the actual key holding the list (e.g. "subscriptions" vs "results")
        subscriptions.extend(data.get("subscriptions", []))

        # VERIFY: pagination shape - could be a "next" URL, a "next_page_token",
        # or offset-based. Adjust this loop accordingly.
        next_token = data.get("next_page_token")
        if next_token:
            params["page_token"] = next_token
            url = f"{BASE_URL}/subscriptions"
        else:
            url = None

        time.sleep(0.2)  # basic rate-limit courtesy

    return subscriptions


def get_contact(contact_id, cache):
    """Fetch a contact's name/email by ID, with a simple cache."""
    if not contact_id:
        return {}
    if contact_id in cache:
        return cache[contact_id]

    url = f"{BASE_URL}/contacts/{contact_id}"
    resp = requests.get(url, headers=HEADERS)
    if resp.status_code != 200:
        cache[contact_id] = {}
        return {}
    contact = resp.json()

    # VERIFY: exact field names for name/email in the contact response
    info = {
        "first_name": contact.get("given_name", ""),
        "last_name": contact.get("family_name", ""),
        "email": next(
            (e.get("email") for e in contact.get("email_addresses", []) if e.get("email")),
            "",
        ),
    }
    cache[contact_id] = info
    time.sleep(0.1)
    return info


def get_product(product_id, cache):
    """Fetch a product's name by ID, with a simple cache."""
    if not product_id:
        return ""
    if product_id in cache:
        return cache[product_id]

    url = f"{BASE_URL}/products/{product_id}"
    resp = requests.get(url, headers=HEADERS)
    if resp.status_code != 200:
        cache[product_id] = ""
        return ""
    product = resp.json()

    # VERIFY: exact field name for product name in the response
    name = product.get("product_name") or product.get("name") or ""
    cache[product_id] = name
    time.sleep(0.1)
    return name


def main():
    subscriptions = get_all_subscriptions()
    print(f"Pulled {len(subscriptions)} subscriptions from the API (all statuses).")

    contact_cache = {}
    product_cache = {}
    rows = []

    for sub in subscriptions:
        contact = get_contact(sub.get("contact_id"), contact_cache)
        product_name = get_product(sub.get("product_id"), product_cache)

        rows.append({
            "subscription_id": sub.get("id"),
            "active": sub.get("active"),
            "contact_id": sub.get("contact_id"),
            "first_name": contact.get("first_name", ""),
            "last_name": contact.get("last_name", ""),
            "email": contact.get("email", ""),
            "product_id": sub.get("product_id"),
            "product_name": product_name,
            "billing_amount": sub.get("billing_amount"),
            "billing_frequency": sub.get("billing_frequency"),
            "billing_cycle": sub.get("billing_cycle"),
            "start_date": sub.get("start_date"),
            "last_bill_date": sub.get("last_bill_date"),
            "next_bill_date": sub.get("next_bill_date"),
            "payment_method_id": sub.get("payment_method_id"),
            "merchant_account_id": sub.get("merchant_account_id"),
            "reason_stopped": sub.get("reason_stopped"),
        })

    with open(OUTPUT_FILE, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [])
        writer.writeheader()
        writer.writerows(rows)

    active_count = sum(1 for r in rows if r["active"])
    print(f"Wrote {len(rows)} subscriptions to {OUTPUT_FILE} ({active_count} marked active).")


if __name__ == "__main__":
    main()