import asyncio
import os
import pandas as pd
from playwright.async_api import async_playwright
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
import requests

# ── Config từ GitHub Secrets ──────────────────────────
PBI_USERNAME  = os.environ["PBI_USERNAME"]
PBI_PASSWORD  = os.environ["PBI_PASSWORD"]
WORKSPACE_ID  = os.environ["WORKSPACE_ID"]
DATASET_ID    = os.environ["DATASET_ID"]
SP_SITE       = os.environ["SHAREPOINT_SITE"]
SP_FOLDER     = os.environ["SHAREPOINT_FOLDER"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]

DAX_QUERY = """
    EVALUATE
    SELECTCOLUMNS(
        MD_QDGP,
        "Segment",       MD_QDGP[Segment],
        "Local_Segment", MD_QDGP[Local_Segment],
        "SKUCode",       MD_QDGP[SKUCode],
        "GroupCode",     MD_QDGP[GroupCode],
        "GroupName",     MD_QDGP[GroupName],
        "Category",      MD_QDGP[Category],
        "Region",        MD_QDGP[Region],
        "MOQ_IT",        MD_QDGP[MOQ (IT)],
        "ApplyForMonth", MD_QDGP[ApplyForMonth],
        "ApplyFrom",     MD_QDGP[ApplyFrom],
        "ApplyTo",       MD_QDGP[ApplyTo]
    )
"""


# ── Login Power BI qua PG PingFederate SSO ───────────
async def get_token_via_browser():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        access_token = None

        async def handle_response(response):
            nonlocal access_token
            if "oauth2/v2.0/token" in response.url and response.status == 200:
                try:
                    body = await response.json()
                    if "access_token" in body:
                        access_token = body["access_token"]
                        print("  ✅ Token captured!")
                except:
                    pass

        page.on("response", handle_response)

        # ── Bước 1: Mở Power BI ──
        print("  Opening Power BI...")
        await page.goto("https://app.powerbi.com")
        await page.wait_for_load_state("networkidle", timeout=30000)

        # ── Bước 2: Nhập email ──
        await page.wait_for_selector('#email', timeout=10000)
        await page.fill('#email', PBI_USERNAME)
        print("  Email filled, submitting...")
        await page.keyboard.press("Enter")

        # ── Bước 3: Chờ sang PG fedauth ──
        await page.wait_for_url("**/fedauth.pg.com/**", timeout=15000)
        await page.wait_for_load_state("networkidle", timeout=15000)
        print(f"  PG SSO: {page.url}")

        # ── Bước 4: Nhập password ──
        await page.wait_for_selector('#password', timeout=10000)
        await page.fill('#password', PBI_PASSWORD)
        print("  Password filled, submitting...")
        await page.keyboard.press("Enter")

        # ── Bước 5: Chờ PingFederate xử lý hết các bước redirect ──
        # PingFederate có nhiều bước intermediate, chờ tổng tối đa 60s
        print("  Waiting for SSO chain to complete...")
        max_wait = 60  # seconds
        interval = 2
        elapsed = 0

        while elapsed < max_wait:
            await page.wait_for_timeout(interval * 1000)
            elapsed += interval
            current_url = page.url
            print(f"  [{elapsed}s] URL: {current_url[:80]}...")

            # Nếu đã về Power BI → dừng chờ
            if "app.powerbi.com" in current_url:
                print("  ✅ Reached Power BI!")
                break

            # Nếu vẫn đang ở trang intermediate của PingFederate
            # → thử click submit nếu có form
            if "fedauth.pg.com" in current_url or "login.microsoftonline.com" in current_url:
                try:
                    submit = page.locator('input[type="submit"], button[type="submit"]')
                    count = await submit.count()
                    if count > 0:
                        await submit.first.click()
                        print(f"  Clicked submit button")
                except:
                    pass

            # Nếu có token rồi thì thoát luôn
            if access_token:
                break

        await page.screenshot(path="step_final.png")
        print(f"  Final URL: {page.url}")

        # Chờ thêm 5s để đảm bảo token được issue
        await page.wait_for_timeout(5000)

        await browser.close()

        if not access_token:
            raise Exception("Could not capture token - check screenshots!")

        return access_token


# ── Fetch data với pagination ─────────────────────────
def fetch_all_rows(token):
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{DATASET_ID}/executeQueries"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    all_rows = []
    page_size = 100000
    start = 0

    while True:
        body = {
            "queries": [{
                "query": DAX_QUERY,
                "pagingInfo": {"start": start, "pageSize": page_size}
            }],
            "serializerSettings": {"includeNulls": True}
        }

        r = requests.post(url, headers=headers, json=body)
        r.raise_for_status()

        rows = r.json()["results"][0]["tables"][0].get("rows", [])

        if not rows:
            print(f"  No more rows at start={start}.")
            break

        all_rows.extend(rows)
        print(f"  Fetched {len(all_rows):,} rows (page start={start})...")

        if len(rows) < page_size:
            break

        start += page_size

    return all_rows


# ── Convert sang CSV ──────────────────────────────────
def rows_to_csv(rows):
    clean_rows = [{k.strip("[]"): v for k, v in row.items()} for row in rows]
    df = pd.DataFrame(clean_rows)
    apply_month = df["ApplyForMonth"].iloc[0] if not df.empty else "unknown"
    filename = f"MD_QDGP_{apply_month}.csv"
    df.to_csv(filename, index=False, encoding="utf-8-sig")
    print(f"  Saved {len(df):,} rows to {filename}")
    return filename, apply_month


# ── Upload lên SharePoint ─────────────────────────────
def upload_to_sharepoint(filename):
    ctx = ClientContext(SP_SITE).with_credentials(
        ClientCredential(CLIENT_ID, CLIENT_SECRET)
    )
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"  Connected to SharePoint: {web.properties['Title']}")

    folder = ctx.web.get_folder_by_server_relative_url(SP_FOLDER)
    with open(filename, "rb") as f:
        folder.upload_file(filename, f.read()).execute_query()
    print(f"  Uploaded '{filename}' to {SP_FOLDER}")


# ── Main ──────────────────────────────────────────────
async def main():
    print("=" * 50)
    print("Power BI Daily Export - MD_QDGP")
    print("=" * 50)

    print("\n[1/4] Logging in via browser...")
    token = await get_token_via_browser()

    print("\n[2/4] Fetching data from Power BI...")
    rows = fetch_all_rows(token)
    print(f"  Total rows: {len(rows):,}")

    if not rows:
        print("  No data. Exiting.")
        exit(1)

    print("\n[3/4] Converting to CSV...")
    filename, month = rows_to_csv(rows)

    print("\n[4/4] Uploading to SharePoint...")
    upload_to_sharepoint(filename)

    print("\n" + "=" * 50)
    print(f"Done! MD_QDGP_{month}.csv uploaded successfully.")
    print("=" * 50)


if __name__ == "__main__":
    asyncio.run(main())
