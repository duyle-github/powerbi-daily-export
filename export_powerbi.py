import requests
import pandas as pd
import os
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

# ── Config từ GitHub Secrets ──────────────────────────
TENANT_ID     = os.environ["TENANT_ID"]
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
WORKSPACE_ID  = os.environ["WORKSPACE_ID"]   # d5309721-c5fc-4d37-924d-0b134d285358
DATASET_ID    = os.environ["DATASET_ID"]     # 271e0b06-2d81-4060-8776-a2ec3f11578a
SP_SITE       = os.environ["SHAREPOINT_SITE"]    # https://pgone.sharepoint.com/sites/ITPG
SP_FOLDER     = os.environ["SHAREPOINT_FOLDER"]  # /Shared Documents/General/Test Automate
PBI_USERNAME  = os.environ["PBI_USERNAME"]   # duy.le@pg.com
PBI_PASSWORD  = os.environ["PBI_PASSWORD"]


# ── Lấy Access Token (dùng Username/Password) ────────
def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "password",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "username": PBI_USERNAME,
        "password": PBI_PASSWORD,
        "scope": "https://analysis.windows.net/powerbi/api/.default"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    token = r.json().get("access_token")
    if not token:
        raise Exception(f"Failed to get token: {r.json()}")
    print("  Token acquired successfully.")
    return token


# ── Fetch data với pagination ─────────────────────────
def fetch_all_rows(token):
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{DATASET_ID}/executeQueries"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

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

    all_rows = []
    page_size = 100000
    start = 0

    while True:
        body = {
            "queries": [{
                "query": DAX_QUERY,
                "pagingInfo": {
                    "start": start,
                    "pageSize": page_size
                }
            }],
            "serializerSettings": {
                "includeNulls": True
            }
        }

        r = requests.post(url, headers=headers, json=body)
        r.raise_for_status()

        result = r.json()
        rows = result["results"][0]["tables"][0].get("rows", [])

        if not rows:
            print(f"  No more rows at start={start}. Done.")
            break

        all_rows.extend(rows)
        print(f"  Fetched {len(all_rows):,} rows so far (page start={start})...")

        # Nếu trả về ít hơn page_size → đã hết data
        if len(rows) < page_size:
            break

        start += page_size

    return all_rows


# ── Convert sang CSV ──────────────────────────────────
def rows_to_csv(rows):
    # Bỏ dấu [] trong tên cột do Power BI API tự thêm vào
    clean_rows = []
    for row in rows:
        clean_rows.append({k.strip("[]"): v for k, v in row.items()})

    df = pd.DataFrame(clean_rows)

    # Lấy ApplyForMonth từ row đầu tiên để đặt tên file
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

    # Kiểm tra kết nối
    web = ctx.web
    ctx.load(web)
    ctx.execute_query()
    print(f"  Connected to SharePoint: {web.properties['Title']}")

    # Upload file (ghi đè nếu đã tồn tại)
    folder = ctx.web.get_folder_by_server_relative_url(SP_FOLDER)
    with open(filename, "rb") as f:
        file_content = f.read()

    folder.upload_file(filename, file_content).execute_query()
    print(f"  Uploaded '{filename}' to {SP_FOLDER}")


# ── Main ──────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 50)
    print("Power BI Daily Export - MD_QDGP")
    print("=" * 50)

    print("\n[1/4] Getting access token...")
    token = get_token()

    print("\n[2/4] Fetching data from Power BI...")
    rows = fetch_all_rows(token)
    print(f"  Total rows fetched: {len(rows):,}")

    if not rows:
        print("  No data returned. Exiting.")
        exit(1)

    print("\n[3/4] Converting to CSV...")
    filename, month = rows_to_csv(rows)
    print(f"  File: {filename}")

    print("\n[4/4] Uploading to SharePoint...")
    upload_to_sharepoint(filename)

    print("\n" + "=" * 50)
    print(f"Done! File MD_QDGP_{month}.csv uploaded successfully.")
    print("=" * 50)
