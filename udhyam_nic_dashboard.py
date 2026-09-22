"""
Udhyam NIC Code Analytics Dashboard — Fancy Edition
=====================================================
Run:  python -m streamlit run udhyam_nic_dashboard.py --server.maxUploadSize 1000

Requirements:
    pip install streamlit pandas openpyxl xlsxwriter plotly boto3

Data source:
    Files are auto-loaded from a Cloudflare R2 bucket (S3-compatible API) using
    boto3. Configure credentials in .streamlit/secrets.toml locally, or in the
    "Secrets" panel of your Streamlit Community Cloud app settings:

        [r2]
        account_id         = "your-cloudflare-account-id"
        access_key_id      = "your-r2-access-key-id"
        secret_access_key  = "your-r2-secret-access-key"
        bucket_name        = "your-bucket-name"
        prefix             = ""   # optional sub-folder inside the bucket

    A manual "Upload files" fallback is still available from the sidebar.
"""

import streamlit as st
import pandas as pd
import json
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import boto3
from botocore.client import Config
from botocore.exceptions import ClientError, EndpointConnectionError

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Udhyam Analytics",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

/* Background */
.stApp { background: #0a0f1e; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #0d1528;
    border-right: 1px solid #1e2d4a;
}
[data-testid="stSidebar"] * { color: #c8d8f0 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stTextInput label,
[data-testid="stSidebar"] .stRadio label,
[data-testid="stSidebar"] .stSlider label { color: #7a9cc8 !important; font-size: 11px !important; text-transform: uppercase; letter-spacing: 0.08em; }

/* Title */
.dash-title {
    font-family: 'Syne', sans-serif;
    font-size: 2.4rem;
    font-weight: 800;
    background: linear-gradient(120deg, #4fc3f7, #81d4fa, #b3e5fc, #e1f5fe);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: -0.02em;
    line-height: 1.1;
    margin-bottom: 0.2rem;
}
.dash-subtitle {
    color: #4a6fa5;
    font-size: 0.85rem;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-bottom: 1.5rem;
}

/* KPI Cards */
.kpi-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 1.5rem; }
.kpi-card {
    background: linear-gradient(135deg, #0d1f3c 0%, #0a1628 100%);
    border: 1px solid #1a3060;
    border-radius: 14px;
    padding: 18px 16px 14px;
    position: relative;
    overflow: hidden;
    transition: transform 0.2s, border-color 0.2s;
}
.kpi-card:hover { transform: translateY(-2px); border-color: #2a5080; }
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 14px 14px 0 0;
}
.kpi-card.blue::before  { background: linear-gradient(90deg, #1565c0, #4fc3f7); }
.kpi-card.teal::before  { background: linear-gradient(90deg, #00695c, #4db6ac); }
.kpi-card.amber::before { background: linear-gradient(90deg, #e65100, #ffb74d); }
.kpi-card.rose::before  { background: linear-gradient(90deg, #880e4f, #f48fb1); }
.kpi-card.indigo::before{ background: linear-gradient(90deg, #283593, #7986cb); }

.kpi-icon { font-size: 1.4rem; margin-bottom: 8px; }
.kpi-label { font-size: 10px; text-transform: uppercase; letter-spacing: 0.12em; color: #4a6fa5; margin-bottom: 4px; }
.kpi-value { font-family: 'Syne', sans-serif; font-size: 1.65rem; font-weight: 700; color: #e8f4fd; line-height: 1; }
.kpi-sub { font-size: 10px; color: #3a5a80; margin-top: 4px; }

/* Section headers */
.section-head {
    font-family: 'Syne', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    color: #7eb8e8;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    border-left: 3px solid #1e6bb8;
    padding-left: 10px;
    margin: 1.5rem 0 0.8rem;
}

/* Chart containers */
.chart-box {
    background: #0d1a30;
    border: 1px solid #1a3060;
    border-radius: 14px;
    padding: 16px;
}

/* Boolean search pills */
.bool-pill {
    display: inline-block;
    background: #0d2040;
    border: 1px solid #1e4080;
    border-radius: 20px;
    padding: 2px 10px;
    font-size: 11px;
    color: #4fc3f7;
    margin: 2px;
}

/* Result banner */
.result-banner {
    background: linear-gradient(135deg, #0a2a1a, #0d3320);
    border: 1px solid #1a6040;
    border-radius: 10px;
    padding: 10px 16px;
    color: #4db88a;
    font-size: 0.9rem;
    margin-bottom: 1rem;
}
.result-banner strong { color: #6dd4a4; }

/* Divider */
.fancy-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #1e3a60, transparent);
    margin: 1.5rem 0;
}

/* Table */
.styled-table { font-size: 12px; }

/* Info tag */
.info-tag {
    display: inline-block;
    background: #0d2040;
    color: #4fc3f7;
    font-size: 10px;
    padding: 2px 8px;
    border-radius: 4px;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)

CHART_THEME = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font_color="#8ab4d8",
    font_family="DM Sans",
)
AXIS_STYLE = dict(
    gridcolor="#0f2040",
    linecolor="#1a3060",
    tickfont_color="#4a6fa5",
    tickfont_size=11,
)
COLOR_SEQ = ["#4fc3f7","#4db6ac","#ffb74d","#f48fb1","#7986cb","#81c784","#ff8a65","#4dd0e1","#ce93d8","#aed581"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_num(n):
    if n >= 1_00_00_000: return f"₹{n/1_00_00_000:.1f} Cr"
    if n >= 1_00_000:    return f"₹{n/1_00_000:.1f} L"
    if n >= 1_000:       return f"₹{n/1_000:.1f} K"
    return f"₹{n:.0f}"

def fmt_int(n):
    if n >= 1_00_00_000: return f"{n/1_00_00_000:.1f} Cr"
    if n >= 1_00_000:    return f"{n/1_00_000:.1f} L"
    if n >= 1_000:       return f"{n/1_000:.1f} K"
    return f"{int(n):,}"

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to Excel bytes for download.

    Uses the xlsxwriter engine (not openpyxl) because certain
    pandas + openpyxl version combinations have a known bug where
    ExcelWriter's internal "remove the default sheet" step can leave the
    workbook with no visible sheet at all, raising
    "IndexError: At least one sheet must be visible" on save. xlsxwriter
    builds the workbook differently and isn't affected by that bug.
    """
    df = df.copy()

    # De-duplicate column names. If the source file already had a column with
    # the same name as one of our aliases (e.g. both "District" and
    # "DISTRICT_NAME"), the alias-rename step can leave two columns with an
    # identical label, which corrupts row.get(...) upstream.
    if df.columns.duplicated().any():
        seen = {}
        new_cols = []
        for c in df.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        df.columns = new_cols

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        # Auto-fit columns (xlsxwriter API: set_column(first_col, last_col, width))
        for i, col_name in enumerate(df.columns):
            if len(df):
                body_max = df[col_name].map(lambda v: len(str(v)) if pd.notna(v) else 0).max()
            else:
                body_max = 0
            max_len = max(body_max, len(str(col_name)))
            ws.set_column(i, i, min(max_len + 4, 60))
    return buf.getvalue()

def xl_btn(df: pd.DataFrame, filename: str, label: str = "⬇️ Export Excel"):
    """Render a small Excel download button below a chart."""
    try:
        data = to_excel_bytes(df)
    except Exception as e:
        st.error(f"❌ Couldn't build this Excel file: {e}")
        return
    st.download_button(
        label=label,
        data=data,
        file_name=f"{filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"xl_{filename}_{id(df)}",
    )

def nic_mask(series, term):
    if not term: return pd.Series([False]*len(series), index=series.index)
    return series == term if len(term) == 5 else series.str.startswith(term)

def get_set(dataframe, term):
    if not term: return set()
    return set(dataframe[nic_mask(dataframe["NIC5DigitId"], term)]["UdyogAadharNo"].unique())

def inv_band(val):
    if pd.isna(val) or val == 0: return "Zero"
    if val <= 25_00_000:         return "Upto ₹25 L (Micro)"
    if val <= 5_00_00_000:       return "₹25 L – ₹5 Cr (Small)"
    if val <= 25_00_00_000:      return "₹5 Cr – ₹25 Cr (Medium)"
    return "Above ₹25 Cr (Large)"

def tur_band(val):
    if pd.isna(val) or val == 0: return "Zero"
    if val <= 5_00_00_000:       return "Upto ₹5 Cr"
    if val <= 50_00_00_000:      return "₹5 Cr – ₹50 Cr"
    if val <= 2_50_00_00_000:    return "₹50 Cr – ₹250 Cr"
    return "Above ₹250 Cr"

# ── Cloudflare R2 (S3-compatible) access ──────────────────────────────────────
R2_EXTENSIONS = (".xlsx", ".xls", ".csv")

@st.cache_resource(show_spinner=False)
def get_r2_client():
    """Build a boto3 S3 client pointed at the Cloudflare R2 endpoint.
    Credentials come from st.secrets["r2"] (see .streamlit/secrets.toml)."""
    cfg = st.secrets["r2"]
    return boto3.client(
        "s3",
        endpoint_url=f"https://{cfg['account_id']}.r2.cloudflarestorage.com",
        aws_access_key_id=cfg["access_key_id"],
        aws_secret_access_key=cfg["secret_access_key"],
        config=Config(signature_version="s3v4"),
        region_name="auto",
    )

@st.cache_data(ttl=300, show_spinner="🔎 Listing files in R2 bucket…")
def list_r2_files(bucket: str, prefix: str = ""):
    """List spreadsheet objects in the bucket/prefix. Cached for 5 minutes."""
    client = get_r2_client()
    paginator = client.get_paginator("list_objects_v2")
    files = []
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix or ""):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            if key.lower().endswith(R2_EXTENSIONS):
                files.append({
                    "key": key,
                    "size": obj["Size"],
                    "last_modified": obj["LastModified"],
                    "etag": obj["ETag"].strip('"'),
                })
    return files

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_r2_object(bucket: str, key: str, etag: str) -> bytes:
    """Download one object's bytes. `etag` is only used as a cache key so that
    a changed file in R2 (new etag) automatically invalidates the cache."""
    client = get_r2_client()
    return client.get_object(Bucket=bucket, Key=key)["Body"].read()

# ── Load & parse ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="⚙️ Parsing NIC codes from your files…")
def load_data(files_bytes):
    all_dfs = []
    for fname, fbytes in files_bytes:
        try:
            buf = BytesIO(fbytes)
            df  = pd.read_csv(buf) if fname.lower().endswith(".csv") else pd.read_excel(buf)
            yr  = "Unknown"
            for y in ["2020","2021","2022","2023","2024","2025","2026"]:
                if y in fname: yr = y; break
            if yr == "Unknown" and "CreateDate" in df.columns:
                try: yr = str(int(pd.to_datetime(df["CreateDate"], errors="coerce").dt.year.mode()[0]))
                except: pass
            df["_year"] = yr
            df["_file"] = fname
            all_dfs.append(df)
        except Exception as e:
            st.warning(f"Could not read {fname}: {e}")
    if not all_dfs: return pd.DataFrame()
    raw = pd.concat(all_dfs, ignore_index=True)

    def parse(val):
        try:
            items = json.loads(val)
            if isinstance(items, list):
                return [(str(i.get("NIC5DigitId","")), i.get("Description","")) for i in items]
        except: pass
        return [("Unknown", "")]

    # ── Normalise column names: strip spaces, fix common variants ──
    raw.columns = raw.columns.str.strip()

    # Map common alternate spellings → canonical names
    col_aliases = {
        "UdyogAadharNo": ["UdyogAadharNo","Udyog Aadhar No","UdyogAadhar No","udyogaadharno","Udyogaadharno"],
        "EnterpriseName": ["EnterpriseName","Enterprise Name","enterprisename"],
        "Gender":         ["Gender","gender"],
        "SocialCategory": ["SocialCategory","Social Category","socialcategory"],
        "Dic_Name":       ["Dic_Name","DIC_Name","DicName","dic_name","DIC Name"],
        "DISTRICT_NAME":  ["DISTRICT_NAME","District_Name","DistrictName","district_name","District Name"],
        "state_name":     ["state_name","State_Name","StateName","State Name"],
        "MajorActivity":  ["MajorActivity","Major Activity","majoractivity"],
        "EnterpriseType": ["EnterpriseType","Enterprise Type","enterprisetype"],
        "OrganisationType":["OrganisationType","Organisation Type"],
        "TotalEmp":       ["TotalEmp","Total Emp","TotalEmployees","total_emp"],
        "InvestmentCost": ["InvestmentCost","Investment Cost","investmentcost"],
        "NetTurnover":    ["NetTurnover","Net Turnover","netturnover"],
        "ActivityDetail": ["ActivityDetail","Activity Detail","activitydetail"],
        "CreateDate":     ["CreateDate","Create Date","createdate"],
    }
    rename_map = {}
    for canonical, aliases in col_aliases.items():
        for alias in aliases:
            if alias in raw.columns and alias != canonical:
                rename_map[alias] = canonical
    if rename_map:
        raw = raw.rename(columns=rename_map)

    # If renaming created a collision (e.g. the file already had both
    # "Dic_Name" and "DIC_Name" as separate columns, which both map to the
    # same canonical name), de-duplicate rather than silently merging them —
    # row.get() on a duplicate-labelled column returns a Series, not a
    # scalar, which corrupts every downstream row and can break Excel export.
    if raw.columns.duplicated().any():
        seen = {}
        new_cols = []
        for c in raw.columns:
            if c in seen:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
            else:
                seen[c] = 0
                new_cols.append(c)
        raw.columns = new_cols

    # Keep every column from the source file (not just a fixed whitelist), so
    # fields like Email, MobileNo, Latitude, Longitude etc. are preserved too.
    # ActivityDetail is dropped since it's already parsed into NIC5DigitId/NICDescription.
    keep = [c for c in raw.columns if c != "ActivityDetail"]

    rows = []
    for _, row in raw.iterrows():
        for nic_id, nic_desc in parse(row.get("ActivityDetail","")):
            e = {c: row.get(c) for c in keep}
            e["NIC5DigitId"] = nic_id
            e["NICDescription"] = nic_desc
            rows.append(e)

    out = pd.DataFrame(rows)
    if "InvestmentCost" in out.columns:
        out["InvestmentCost"] = pd.to_numeric(out["InvestmentCost"], errors="coerce").fillna(0)
        out["InvBand"] = out["InvestmentCost"].apply(inv_band)
    if "NetTurnover" in out.columns:
        out["NetTurnover"] = pd.to_numeric(out["NetTurnover"], errors="coerce").fillna(0)
        out["TurBand"] = out["NetTurnover"].apply(tur_band)
    if "TotalEmp" in out.columns:
        out["TotalEmp"] = pd.to_numeric(out["TotalEmp"], errors="coerce").fillna(0)
    return out

# ── Sidebar ───────────────────────────────────────────────────────────────────
file_data = []  # list of (filename, bytes) — filled from R2 or manual upload

with st.sidebar:
    st.markdown("### 🏭 Udhyam Analytics")
    st.markdown("---")
    st.markdown("#### ☁️ Data Source")

    try:
        has_r2_secrets = "r2" in st.secrets
    except Exception:
        has_r2_secrets = False
    data_source = st.radio(
        "Load data from",
        ["Cloudflare R2 (auto)", "Upload files manually"],
        index=0 if has_r2_secrets else 1,
        horizontal=False,
    )

    if data_source == "Upload files manually":
        uploaded = st.file_uploader(
            "Upload Excel / CSV files",
            type=["xlsx","xls","csv"],
            accept_multiple_files=True,
            help="Upload one or more year-wise Udhyam data files"
        )
        if uploaded:
            file_data = [(f.name, f.read()) for f in uploaded]

    else:
        if not has_r2_secrets:
            st.error(
                "No R2 credentials found. Add an **[r2]** section to "
                "`.streamlit/secrets.toml` (locally) or to your app's **Secrets** "
                "panel on Streamlit Community Cloud — see the script's docstring "
                "for the exact format."
            )
        else:
            cfg = st.secrets["r2"]
            try:
                bucket = cfg["bucket_name"]
            except KeyError:
                st.error("Your [r2] secrets are missing `bucket_name`.")
                bucket = None

            if bucket:
                default_prefix = cfg.get("prefix", "")
                prefix = st.text_input("Folder / prefix (optional)", value=default_prefix)

                col_refresh, col_caption = st.columns([1, 4])
                with col_refresh:
                    if st.button("🔄", help="Refresh file list from R2"):
                        list_r2_files.clear()
                        st.rerun()

                try:
                    r2_files = list_r2_files(bucket, prefix)
                except (ClientError, EndpointConnectionError, Exception) as e:
                    st.error(f"❌ Could not connect to R2: {e}")
                    r2_files = []

                if not r2_files:
                    st.warning("No .xlsx / .xls / .csv files found in this bucket/prefix.")
                else:
                    file_keys = [f["key"] for f in r2_files]
                    sel_keys = st.multiselect(
                        "Files to include", file_keys, default=file_keys,
                        help="All spreadsheet files found in the bucket are selected by default.",
                    )
                    for f in r2_files:
                        if f["key"] in sel_keys:
                            try:
                                fbytes = fetch_r2_object(bucket, f["key"], f["etag"])
                                file_data.append((f["key"].split("/")[-1], fbytes))
                            except Exception as e:
                                st.warning(f"Could not fetch `{f['key']}`: {e}")
                    st.caption(f"📦 {len(file_data)} of {len(r2_files)} file(s) loaded from R2")

    if file_data:
        st.markdown("---")
        st.markdown("#### ⚙️ Filters")

    st.markdown("---")
    st.markdown("#### 🔎 NIC Boolean Search")
    st.caption("3-digit prefix  OR  full 5-digit code")
    nic1 = st.text_input("NIC Code 1", placeholder="e.g. 131 or 13911", max_chars=5).strip()
    bool_op = st.radio("Operator", ["AND","OR","NOT"], horizontal=True)
    nic2 = st.text_input("NIC Code 2 (optional)", placeholder="e.g. 139 or 13136", max_chars=5).strip()
    st.caption("""
**AND** → both codes present  
**OR** → either code present  
**NOT** → Code 1 but not Code 2
    """)

if not file_data:
    st.markdown('<div class="dash-title">Udhyam NIC Analytics</div>', unsafe_allow_html=True)
    st.markdown('<div class="dash-subtitle">MSME Enterprise Intelligence Dashboard</div>', unsafe_allow_html=True)
    st.info("👈  Select a data source from the sidebar to begin (Cloudflare R2 or manual upload).")
    st.stop()

# ── Load data ─────────────────────────────────────────────────────────────────
df = load_data(file_data)
if df.empty:
    st.error("No data could be parsed. Please check your files."); st.stop()

if "UdyogAadharNo" not in df.columns:
    st.error(
        f"❌ Could not find **UdyogAadharNo** column.\n\n"
        f"Columns found in your file: `{', '.join(df.columns.tolist())}`\n\n"
        "Please check that your file has a column named **UdyogAadharNo**."
    )
    st.stop()

if "UdyogAadharNo" not in df.columns:
    st.error(
        f"❌ Could not find **UdyogAadharNo** column in your files.\n\n"
        f"Columns found: `{', '.join(df.columns.tolist())}`\n\n"
        f"Please check your file has a column named exactly **UdyogAadharNo** (or a known variant)."
    )
    st.stop()

# ── Sidebar filters (dynamic) ─────────────────────────────────────────────────
with st.sidebar:
    if file_data:
        years = sorted(df["_year"].dropna().unique())
        sel_years = st.multiselect("Year", years, default=years)

        if "MajorActivity" in df.columns:
            acts = sorted(df["MajorActivity"].dropna().unique())
            sel_act = st.multiselect("Major Activity", acts, default=acts)
        else: sel_act = None

        if "Dic_Name" in df.columns:
            dics = sorted(df["Dic_Name"].dropna().unique())
            sel_dic = st.multiselect("DIC / District", dics, default=dics)
        else: sel_dic = None

        if "EnterpriseType" in df.columns:
            etypes = sorted(df["EnterpriseType"].dropna().unique())
            sel_etype = st.multiselect("Enterprise Type", etypes, default=etypes)
        else: sel_etype = None

        if "Gender" in df.columns:
            gens = sorted(df["Gender"].dropna().unique())
            sel_gen = st.multiselect("Gender", gens, default=gens)
        else: sel_gen = None

        if "SocialCategory" in df.columns:
            cats = sorted(df["SocialCategory"].dropna().unique())
            sel_cat = st.multiselect("Social Category", cats, default=cats)
        else: sel_cat = None

        top_n = st.slider("Top N for charts", 5, 30, 12)

# ── Apply filters ─────────────────────────────────────────────────────────────
mask = df["_year"].isin(sel_years)
if sel_act  and "MajorActivity"  in df.columns: mask &= df["MajorActivity"].isin(sel_act)
if sel_dic  and "Dic_Name"       in df.columns: mask &= df["Dic_Name"].isin(sel_dic)
if sel_etype and "EnterpriseType" in df.columns: mask &= df["EnterpriseType"].isin(sel_etype)
if sel_gen  and "Gender"         in df.columns: mask &= df["Gender"].isin(sel_gen)
if sel_cat  and "SocialCategory" in df.columns: mask &= df["SocialCategory"].isin(sel_cat)
fdf = df[mask].copy()

# ── Boolean NIC filter ────────────────────────────────────────────────────────
bool_fdf = fdf.copy()
bool_label = ""
if nic1:
    s1 = get_set(fdf, nic1)
    s2 = get_set(fdf, nic2) if nic2 else set()
    if nic2:
        if bool_op == "AND":  matched = s1 & s2; bool_label = f"{nic1} AND {nic2}"
        elif bool_op == "OR": matched = s1 | s2; bool_label = f"{nic1} OR {nic2}"
        else:                 matched = s1 - s2; bool_label = f"{nic1} NOT {nic2}"
    else:
        matched = s1; bool_label = nic1
    bool_fdf = fdf[fdf["UdyogAadharNo"].isin(matched)]

# ── Build nic_count from bool_fdf ────────────────────────────────────────────
nic_count = (
    bool_fdf.groupby(["NIC5DigitId","NICDescription"])["UdyogAadharNo"]
    .nunique().reset_index()
    .rename(columns={"UdyogAadharNo":"EnterpriseCount"})
    .sort_values("EnterpriseCount", ascending=False)
)

# ── NIC filter for charts ─────────────────────────────────────────────────────
if nic1:
    nic_m = nic_mask(nic_count["NIC5DigitId"], nic1)
    if nic2: nic_m = nic_m | nic_mask(nic_count["NIC5DigitId"], nic2)
    nic_chart_df = nic_count[nic_m] if nic_m.any() else nic_count.head(top_n)
    chart_title = f"NIC  {bool_label}"
else:
    nic_chart_df = nic_count.head(top_n)
    chart_title = f"Top {top_n} NIC Codes"

# ── Derived metrics ───────────────────────────────────────────────────────────
uniq_ent    = bool_fdf["UdyogAadharNo"].nunique()
uniq_nic    = bool_fdf["NIC5DigitId"].nunique()
total_emp   = int(bool_fdf.drop_duplicates("UdyogAadharNo")["TotalEmp"].sum()) if "TotalEmp" in bool_fdf.columns else 0
total_inv   = bool_fdf.drop_duplicates("UdyogAadharNo")["InvestmentCost"].sum() if "InvestmentCost" in bool_fdf.columns else 0
total_tur   = bool_fdf.drop_duplicates("UdyogAadharNo")["NetTurnover"].sum() if "NetTurnover" in bool_fdf.columns else 0

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown('<div class="dash-title">Udhyam NIC Analytics</div>', unsafe_allow_html=True)
st.markdown('<div class="dash-subtitle">MSME Enterprise Intelligence · Gujarat</div>', unsafe_allow_html=True)

if bool_label:
    st.markdown(f"""<div class="result-banner">
        🔎 Boolean Filter Active — <strong>{bool_label}</strong> &nbsp;|&nbsp;
        <strong>{uniq_ent:,}</strong> enterprises matched
    </div>""", unsafe_allow_html=True)

# ── KPI CARDS ─────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="kpi-grid">
  <div class="kpi-card blue">
    <div class="kpi-icon">🏭</div>
    <div class="kpi-label">Enterprises</div>
    <div class="kpi-value">{fmt_int(uniq_ent)}</div>
    <div class="kpi-sub">Unique UdyogAadharNo</div>
  </div>
  <div class="kpi-card teal">
    <div class="kpi-icon">🗂️</div>
    <div class="kpi-label">NIC Codes</div>
    <div class="kpi-value">{uniq_nic:,}</div>
    <div class="kpi-sub">Unique 5-digit codes</div>
  </div>
  <div class="kpi-card amber">
    <div class="kpi-icon">👷</div>
    <div class="kpi-label">Total Employees</div>
    <div class="kpi-value">{fmt_int(total_emp)}</div>
    <div class="kpi-sub">Sum of TotalEmp</div>
  </div>
  <div class="kpi-card rose">
    <div class="kpi-icon">💰</div>
    <div class="kpi-label">Investment</div>
    <div class="kpi-value">{fmt_num(total_inv)}</div>
    <div class="kpi-sub">Total InvestmentCost</div>
  </div>
  <div class="kpi-card indigo">
    <div class="kpi-icon">📈</div>
    <div class="kpi-label">Net Turnover</div>
    <div class="kpi-value">{fmt_num(total_tur)}</div>
    <div class="kpi-sub">Total NetTurnover</div>
  </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── FULL RAW DATA EXPORT (user-chosen columns) ───────────────────────────────
st.markdown('<div class="section-head">📤 Export Filtered Data</div>', unsafe_allow_html=True)
st.caption("Export the actual filtered records (not just chart summaries) — pick exactly the columns you need.")

exp_c1, exp_c2 = st.columns([1, 2])

with exp_c1:
    export_grain = st.radio(
        "Row granularity",
        ["One row per enterprise (deduplicated)", "One row per enterprise × NIC code (detailed)"],
        key="export_grain",
    )

# Columns available on the currently filtered data (excludes internal helper "_file")
all_export_cols = [c for c in bool_fdf.columns if c != "_file"]

# Sensible default selection (only columns that actually exist)
_preferred_defaults = [
    "UdyogAadharNo", "EnterpriseName", "Gender", "SocialCategory",
    "Dic_Name", "DISTRICT_NAME", "state_name", "MajorActivity",
    "EnterpriseType", "OrganisationType", "TotalEmp",
    "InvestmentCost", "InvBand", "NetTurnover", "TurBand",
    "CreateDate", "_year", "NIC5DigitId", "NICDescription",
]
default_export_cols = [c for c in _preferred_defaults if c in all_export_cols] or all_export_cols

with exp_c2:
    sel_export_cols = st.multiselect(
        "Columns to include",
        options=all_export_cols,
        default=default_export_cols,
        key="export_cols",
        help="Only the columns you select here will appear in the exported Excel file.",
    )

if sel_export_cols:
    if export_grain.startswith("One row per enterprise ("):
        export_df = bool_fdf.drop_duplicates("UdyogAadharNo")[sel_export_cols].copy()
    else:
        export_df = bool_fdf[sel_export_cols].copy()

    st.markdown(
        f'<div class="result-banner">📋 <strong>{len(export_df):,}</strong> rows × '
        f'<strong>{len(sel_export_cols)}</strong> columns ready to export'
        f'{f" — filtered to <strong>{bool_label}</strong>" if bool_label else ""}</div>',
        unsafe_allow_html=True,
    )
    xl_btn(export_df, "udhyam_filtered_export", "⬇️  Download Filtered Data (Excel)")
else:
    st.info("Select at least one column above to enable the export.")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 1: NIC bar + Major Activity donut ────────────────────────────────────
st.markdown('<div class="section-head">NIC Code Distribution</div>', unsafe_allow_html=True)
c1, c2 = st.columns([3, 1])

with c1:
    nic_chart_df["Label"] = nic_chart_df["NIC5DigitId"] + "  " + nic_chart_df["NICDescription"].str[:45]
    fig = go.Figure(go.Bar(
        x=nic_chart_df["EnterpriseCount"],
        y=nic_chart_df["Label"],
        orientation="h",
        marker=dict(
            color=nic_chart_df["EnterpriseCount"],
            colorscale=[[0,"#0d2a4a"],[0.5,"#1565c0"],[1,"#4fc3f7"]],
            showscale=False,
        ),
        text=nic_chart_df["EnterpriseCount"],
        textposition="outside",
        textfont=dict(color="#7eb8e8", size=11),
    ))
    fig.update_layout(
        **CHART_THEME,
        title=dict(text=chart_title, font=dict(family="Syne", size=13, color="#7eb8e8")),
        yaxis=dict(autorange="reversed", **AXIS_STYLE),
        xaxis=dict(title="Enterprises", **AXIS_STYLE),
        margin=dict(l=10, r=30, t=40, b=10),
        height=max(320, len(nic_chart_df)*30 + 60),
    )
    st.plotly_chart(fig, use_container_width=True)
    xl_btn(nic_chart_df[["NIC5DigitId","NICDescription","EnterpriseCount"]], "nic_distribution")

with c2:
    if "MajorActivity" in bool_fdf.columns:
        act_df = (
            bool_fdf.groupby("MajorActivity")["UdyogAadharNo"].nunique()
            .reset_index().rename(columns={"UdyogAadharNo":"Count"})
        )
        fig2 = go.Figure(go.Pie(
            labels=act_df["MajorActivity"],
            values=act_df["Count"],
            hole=0.55,
            marker=dict(colors=["#4fc3f7","#ffb74d","#4db6ac"], line=dict(color="#0a0f1e", width=2)),
            textinfo="label+percent",
            textfont=dict(size=11, color="#c8d8f0"),
        ))
        fig2.update_layout(
            **CHART_THEME,
            title=dict(text="Major Activity", font=dict(family="Syne", size=13, color="#7eb8e8")),
            showlegend=False,
            margin=dict(l=10, r=10, t=40, b=10),
            height=320,
            annotations=[dict(text=f"<b>{uniq_ent:,}</b>", x=0.5, y=0.5,
                              font=dict(size=18, color="#e8f4fd", family="Syne"),
                              showarrow=False)]
        )
        st.plotly_chart(fig2, use_container_width=True)
        xl_btn(act_df.rename(columns={"Count":"Enterprises"}), "major_activity")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 2: DIC/District bar + Enterprise Type bar ─────────────────────────────
st.markdown('<div class="section-head">District & Enterprise Type</div>', unsafe_allow_html=True)
c3, c4 = st.columns(2)

with c3:
    dist_col = "Dic_Name" if "Dic_Name" in bool_fdf.columns else "DISTRICT_NAME"
    if dist_col in bool_fdf.columns:
        dist_df = (
            bool_fdf.groupby(dist_col)["UdyogAadharNo"].nunique()
            .reset_index().rename(columns={"UdyogAadharNo":"Count", dist_col:"District"})
            .sort_values("Count", ascending=True).tail(top_n)
        )
        emp_dist = (
            bool_fdf.drop_duplicates("UdyogAadharNo")
            .groupby(dist_col)["TotalEmp"].sum()
            .reset_index().rename(columns={dist_col:"District"})
        ) if "TotalEmp" in bool_fdf.columns else None

        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            y=dist_df["District"], x=dist_df["Count"],
            orientation="h", name="Enterprises",
            marker=dict(color="#1565c0"),
            text=dist_df["Count"], textposition="outside",
            textfont=dict(color="#7eb8e8", size=10),
        ))
        fig3.update_layout(
            **CHART_THEME,
            title=dict(text=f"Enterprises by DIC District (Top {top_n})", font=dict(family="Syne", size=13, color="#7eb8e8")),
            xaxis=dict(title="Enterprises", **AXIS_STYLE),
            yaxis=dict(**AXIS_STYLE),
            margin=dict(l=10, r=30, t=40, b=10),
            height=380, showlegend=False,
        )
        st.plotly_chart(fig3, use_container_width=True)
        xl_btn(dist_df, "district_enterprises")

with c4:
    if "EnterpriseType" in bool_fdf.columns:
        et_df = (
            bool_fdf.groupby("EnterpriseType")["UdyogAadharNo"].nunique()
            .reset_index().rename(columns={"UdyogAadharNo":"Count"})
            .sort_values("Count", ascending=False)
        )
        emp_et = (
            bool_fdf.drop_duplicates("UdyogAadharNo")
            .groupby("EnterpriseType")["TotalEmp"].sum()
            .reset_index()
        ) if "TotalEmp" in bool_fdf.columns else None

        fig4 = go.Figure()
        fig4.add_trace(go.Bar(
            x=et_df["EnterpriseType"], y=et_df["Count"],
            marker=dict(color=COLOR_SEQ[:len(et_df)]),
            text=et_df["Count"], textposition="outside",
            textfont=dict(color="#7eb8e8", size=12),
        ))
        if emp_et is not None:
            merged = et_df.merge(emp_et, on="EnterpriseType")
            fig4.add_trace(go.Scatter(
                x=merged["EnterpriseType"], y=merged["TotalEmp"],
                mode="lines+markers+text",
                name="Employees",
                yaxis="y2",
                line=dict(color="#ffb74d", width=2),
                marker=dict(size=8, color="#ffb74d"),
                text=merged["TotalEmp"].apply(lambda v: fmt_int(v)),
                textposition="top center",
                textfont=dict(color="#ffb74d", size=10),
            ))
            fig4.update_layout(yaxis2=dict(
                overlaying="y", side="right",
                title="Employees", **AXIS_STYLE,
                title_font=dict(color="#ffb74d"),
            ))

        fig4.update_layout(
            **CHART_THEME,
            title=dict(text="Enterprise Type — Count & Employees", font=dict(family="Syne", size=13, color="#7eb8e8")),
            xaxis=dict(**AXIS_STYLE),
            yaxis=dict(title="Enterprises", **AXIS_STYLE),
            margin=dict(l=10, r=60, t=40, b=10),
            height=380, showlegend=True,
            legend=dict(font=dict(color="#8ab4d8", size=10), bgcolor="rgba(0,0,0,0)"),
        )
        st.plotly_chart(fig4, use_container_width=True)
        et_export = et_df.merge(emp_et, on="EnterpriseType", how="left") if emp_et is not None else et_df
        xl_btn(et_export.rename(columns={"Count":"Enterprises","TotalEmp":"Employees"}), "enterprise_type")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 3: Gender & Social Category × Enterprise Type ────────────────────────
st.markdown('<div class="section-head">Gender & Social Category by Enterprise Type</div>', unsafe_allow_html=True)

ET_COLORS = {"Micro": "#4fc3f7", "Small": "#ffb74d", "Medium": "#4db6ac"}
ET_ORDER   = ["Micro", "Small", "Medium"]

# ── Gender × Enterprise Type ──────────────────────────────────────────────────
if "Gender" in bool_fdf.columns and "EnterpriseType" in bool_fdf.columns:
    st.markdown("##### 👤 Gender-wise")
    gen_et = (
        bool_fdf.groupby(["Gender", "EnterpriseType"])["UdyogAadharNo"].nunique()
        .reset_index().rename(columns={"UdyogAadharNo": "Count"})
    )
    # Overall gender totals for the summary donut
    gen_total = gen_et.groupby("Gender")["Count"].sum().reset_index()

    ga, gb, gc = st.columns([1.2, 2, 2])

    with ga:
        # Summary donut — total gender split
        fig_gen_donut = go.Figure(go.Pie(
            labels=gen_total["Gender"], values=gen_total["Count"],
            hole=0.55,
            marker=dict(colors=["#4fc3f7","#f48fb1","#4db6ac"],
                        line=dict(color="#0a0f1e", width=2)),
            textinfo="label+percent",
            textfont=dict(size=11, color="#c8d8f0"),
        ))
        fig_gen_donut.update_layout(
            **CHART_THEME,
            title=dict(text="Overall Split", font=dict(family="Syne", size=12, color="#7eb8e8")),
            showlegend=False, margin=dict(l=5,r=5,t=36,b=5), height=240,
            annotations=[dict(text=f"<b>{gen_total['Count'].sum():,}</b>",
                              x=0.5, y=0.5, showarrow=False,
                              font=dict(size=15, color="#e8f4fd", family="Syne"))]
        )
        st.plotly_chart(fig_gen_donut, use_container_width=True)
        xl_btn(gen_total.rename(columns={"Count":"Enterprises"}), "gender_split")

    with gb:
        # Grouped bar — gender × enterprise type
        fig_gen_bar = go.Figure()
        et_present = [e for e in ET_ORDER if e in gen_et["EnterpriseType"].values]
        for et in et_present:
            sub = gen_et[gen_et["EnterpriseType"] == et]
            fig_gen_bar.add_trace(go.Bar(
                name=et, x=sub["Gender"], y=sub["Count"],
                marker_color=ET_COLORS.get(et, "#7986cb"),
                text=sub["Count"], textposition="outside",
                textfont=dict(size=11, color="#c8d8f0"),
            ))
        fig_gen_bar.update_layout(
            **CHART_THEME,
            title=dict(text="Gender × Enterprise Type", font=dict(family="Syne", size=12, color="#7eb8e8")),
            barmode="group",
            xaxis=dict(**AXIS_STYLE),
            yaxis=dict(title="Enterprises", **AXIS_STYLE),
            legend=dict(font=dict(color="#8ab4d8", size=10), bgcolor="rgba(0,0,0,0)",
                        orientation="h", y=1.12),
            margin=dict(l=10,r=10,t=44,b=10), height=240,
        )
        st.plotly_chart(fig_gen_bar, use_container_width=True)
        xl_btn(gen_et.rename(columns={"Count":"Enterprises"}), "gender_x_enterprise_type")

    with gc:
        # Stacked % bar — enterprise type breakdown within each gender
        fig_gen_pct = go.Figure()
        for et in et_present:
            sub = gen_et[gen_et["EnterpriseType"] == et]
            fig_gen_pct.add_trace(go.Bar(
                name=et, x=sub["Gender"], y=sub["Count"],
                marker_color=ET_COLORS.get(et, "#7986cb"),
                text=sub["Count"].apply(lambda v: f"{v}"),
                textposition="inside",
                textfont=dict(size=10, color="#0a0f1e"),
            ))
        fig_gen_pct.update_layout(
            **CHART_THEME,
            title=dict(text="Gender — Stacked by Type", font=dict(family="Syne", size=12, color="#7eb8e8")),
            barmode="stack",
            xaxis=dict(**AXIS_STYLE),
            yaxis=dict(title="Enterprises", **AXIS_STYLE),
            legend=dict(font=dict(color="#8ab4d8", size=10), bgcolor="rgba(0,0,0,0)",
                        orientation="h", y=1.12),
            margin=dict(l=10,r=10,t=44,b=10), height=240,
        )
        st.plotly_chart(fig_gen_pct, use_container_width=True)
        xl_btn(gen_et.rename(columns={"Count":"Enterprises"}), "gender_stacked")

    # Mini summary table for gender × enterprise type
    gen_pivot = gen_et.pivot_table(
        index="Gender", columns="EnterpriseType", values="Count", aggfunc="sum", fill_value=0
    ).reset_index()
    gen_pivot["Total"] = gen_pivot.select_dtypes("number").sum(axis=1)
    gen_pivot = gen_pivot.sort_values("Total", ascending=False)
    st.dataframe(gen_pivot, use_container_width=True, hide_index=True, height=130)
    xl_btn(gen_pivot, "gender_enterprise_summary")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── Social Category × Enterprise Type ────────────────────────────────────────
if "SocialCategory" in bool_fdf.columns and "EnterpriseType" in bool_fdf.columns:
    st.markdown("##### 🏷️ Social Category-wise")
    sc_et = (
        bool_fdf.groupby(["SocialCategory", "EnterpriseType"])["UdyogAadharNo"].nunique()
        .reset_index().rename(columns={"UdyogAadharNo": "Count"})
    )
    sc_total = sc_et.groupby("SocialCategory")["Count"].sum().reset_index().sort_values("Count", ascending=False)

    sa, sb, sc_col = st.columns([1.2, 2, 2])

    with sa:
        SC_PALETTE = ["#7986cb","#ffb74d","#81c784","#f48fb1","#4dd0e1"]
        fig_sc_donut = go.Figure(go.Pie(
            labels=sc_total["SocialCategory"], values=sc_total["Count"],
            hole=0.55,
            marker=dict(colors=SC_PALETTE[:len(sc_total)],
                        line=dict(color="#0a0f1e", width=2)),
            textinfo="label+percent",
            textfont=dict(size=11, color="#c8d8f0"),
        ))
        fig_sc_donut.update_layout(
            **CHART_THEME,
            title=dict(text="Overall Split", font=dict(family="Syne", size=12, color="#7eb8e8")),
            showlegend=False, margin=dict(l=5,r=5,t=36,b=5), height=260,
            annotations=[dict(text=f"<b>{sc_total['Count'].sum():,}</b>",
                              x=0.5, y=0.5, showarrow=False,
                              font=dict(size=15, color="#e8f4fd", family="Syne"))]
        )
        st.plotly_chart(fig_sc_donut, use_container_width=True)
        xl_btn(sc_total.rename(columns={"Count":"Enterprises"}), "social_category_split")

    with sb:
        fig_sc_bar = go.Figure()
        et_present = [e for e in ET_ORDER if e in sc_et["EnterpriseType"].values]
        for et in et_present:
            sub = sc_et[sc_et["EnterpriseType"] == et]
            fig_sc_bar.add_trace(go.Bar(
                name=et, x=sub["SocialCategory"], y=sub["Count"],
                marker_color=ET_COLORS.get(et, "#7986cb"),
                text=sub["Count"], textposition="outside",
                textfont=dict(size=11, color="#c8d8f0"),
            ))
        fig_sc_bar.update_layout(
            **CHART_THEME,
            title=dict(text="Social Category × Enterprise Type", font=dict(family="Syne", size=12, color="#7eb8e8")),
            barmode="group",
            xaxis=dict(**AXIS_STYLE),
            yaxis=dict(title="Enterprises", **AXIS_STYLE),
            legend=dict(font=dict(color="#8ab4d8", size=10), bgcolor="rgba(0,0,0,0)",
                        orientation="h", y=1.12),
            margin=dict(l=10,r=10,t=44,b=10), height=260,
        )
        st.plotly_chart(fig_sc_bar, use_container_width=True)
        xl_btn(sc_et.rename(columns={"Count":"Enterprises"}), "social_category_x_enterprise_type")

    with sc_col:
        fig_sc_stack = go.Figure()
        for et in et_present:
            sub = sc_et[sc_et["EnterpriseType"] == et]
            fig_sc_stack.add_trace(go.Bar(
                name=et, x=sub["SocialCategory"], y=sub["Count"],
                marker_color=ET_COLORS.get(et, "#7986cb"),
                text=sub["Count"].apply(lambda v: f"{v}"),
                textposition="inside",
                textfont=dict(size=10, color="#0a0f1e"),
            ))
        fig_sc_stack.update_layout(
            **CHART_THEME,
            title=dict(text="Social Category — Stacked by Type", font=dict(family="Syne", size=12, color="#7eb8e8")),
            barmode="stack",
            xaxis=dict(**AXIS_STYLE),
            yaxis=dict(title="Enterprises", **AXIS_STYLE),
            legend=dict(font=dict(color="#8ab4d8", size=10), bgcolor="rgba(0,0,0,0)",
                        orientation="h", y=1.12),
            margin=dict(l=10,r=10,t=44,b=10), height=260,
        )
        st.plotly_chart(fig_sc_stack, use_container_width=True)
        xl_btn(sc_et.rename(columns={"Count":"Enterprises"}), "social_category_stacked")

    # Mini summary table for social category × enterprise type
    sc_pivot = sc_et.pivot_table(
        index="SocialCategory", columns="EnterpriseType", values="Count", aggfunc="sum", fill_value=0
    ).reset_index()
    sc_pivot["Total"] = sc_pivot.select_dtypes("number").sum(axis=1)
    sc_pivot = sc_pivot.sort_values("Total", ascending=False)
    st.dataframe(sc_pivot, use_container_width=True, hide_index=True, height=160)
    xl_btn(sc_pivot, "social_category_enterprise_summary")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 4: Investment bands + Turnover bands ──────────────────────────────────
st.markdown('<div class="section-head">Investment Cost & Net Turnover Bands</div>', unsafe_allow_html=True)
c7, c8 = st.columns(2)

INV_ORDER = ["Zero","Upto ₹25 L (Micro)","₹25 L – ₹5 Cr (Small)","₹5 Cr – ₹25 Cr (Medium)","Above ₹25 Cr (Large)"]
TUR_ORDER = ["Zero","Upto ₹5 Cr","₹5 Cr – ₹50 Cr","₹50 Cr – ₹250 Cr","Above ₹250 Cr"]

with c7:
    if "InvBand" in bool_fdf.columns:
        inv_df = (
            bool_fdf.groupby("InvBand")["UdyogAadharNo"].nunique()
            .reset_index().rename(columns={"UdyogAadharNo":"Count"})
        )
        inv_df["InvBand"] = pd.Categorical(inv_df["InvBand"], categories=INV_ORDER, ordered=True)
        inv_df = inv_df.sort_values("InvBand")
        fig7 = go.Figure(go.Bar(
            x=inv_df["InvBand"], y=inv_df["Count"],
            marker=dict(color=["#263238","#1565c0","#1976d2","#1e88e5","#42a5f5"][:len(inv_df)]),
            text=inv_df["Count"], textposition="outside",
            textfont=dict(color="#7eb8e8", size=12),
        ))
        fig7.update_layout(
            **CHART_THEME,
            title=dict(text="Enterprises by Investment Cost Band", font=dict(family="Syne", size=13, color="#7eb8e8")),
            xaxis=dict(**AXIS_STYLE, tickangle=-20),
            yaxis=dict(title="No. of Enterprises", **AXIS_STYLE),
            margin=dict(l=10,r=10,t=40,b=60), height=320, showlegend=False,
        )
        st.plotly_chart(fig7, use_container_width=True)
        xl_btn(inv_df.rename(columns={"InvBand":"InvestmentBand","Count":"Enterprises"}), "investment_bands")

with c8:
    if "TurBand" in bool_fdf.columns:
        tur_df = (
            bool_fdf.groupby("TurBand")["UdyogAadharNo"].nunique()
            .reset_index().rename(columns={"UdyogAadharNo":"Count"})
        )
        tur_df["TurBand"] = pd.Categorical(tur_df["TurBand"], categories=TUR_ORDER, ordered=True)
        tur_df = tur_df.sort_values("TurBand")
        fig8 = go.Figure(go.Bar(
            x=tur_df["TurBand"], y=tur_df["Count"],
            marker=dict(color=["#263238","#00695c","#00897b","#26a69a","#4db6ac"][:len(tur_df)]),
            text=tur_df["Count"], textposition="outside",
            textfont=dict(color="#7eb8e8", size=12),
        ))
        fig8.update_layout(
            **CHART_THEME,
            title=dict(text="Enterprises by Net Turnover Band", font=dict(family="Syne", size=13, color="#7eb8e8")),
            xaxis=dict(**AXIS_STYLE, tickangle=-20),
            yaxis=dict(title="No. of Enterprises", **AXIS_STYLE),
            margin=dict(l=10,r=10,t=40,b=60), height=320, showlegend=False,
        )
        st.plotly_chart(fig8, use_container_width=True)
        xl_btn(tur_df.rename(columns={"TurBand":"TurnoverBand","Count":"Enterprises"}), "turnover_bands")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 5: Year-wise trend ───────────────────────────────────────────────────
st.markdown('<div class="section-head">Year-wise Registration Trend</div>', unsafe_allow_html=True)
year_df = (
    bool_fdf.groupby("_year")["UdyogAadharNo"].nunique()
    .reset_index().rename(columns={"UdyogAadharNo":"Count","_year":"Year"})
    .sort_values("Year")
)
emp_year = (
    bool_fdf.drop_duplicates("UdyogAadharNo")
    .groupby("_year")["TotalEmp"].sum()
    .reset_index().rename(columns={"_year":"Year"})
) if "TotalEmp" in bool_fdf.columns else None

fig9 = go.Figure()
fig9.add_trace(go.Bar(
    x=year_df["Year"], y=year_df["Count"],
    name="Enterprises", marker=dict(color="#1565c0", opacity=0.85),
    text=year_df["Count"], textposition="outside",
    textfont=dict(color="#7eb8e8", size=12),
))
if emp_year is not None:
    merged_yr = year_df.merge(emp_year, on="Year", how="left")
    fig9.add_trace(go.Scatter(
        x=merged_yr["Year"], y=merged_yr["TotalEmp"],
        mode="lines+markers+text", name="Employees",
        yaxis="y2", line=dict(color="#ffb74d", width=2.5),
        marker=dict(size=9, color="#ffb74d"),
        text=merged_yr["TotalEmp"].apply(lambda v: fmt_int(v) if pd.notna(v) else ""),
        textposition="top center", textfont=dict(color="#ffb74d", size=10),
    ))
    fig9.update_layout(yaxis2=dict(
        overlaying="y", side="right",
        title="Employees", **AXIS_STYLE,
        title_font=dict(color="#ffb74d"),
    ))
fig9.update_layout(
    **CHART_THEME,
    title=dict(text="Year-wise Enterprises & Employees", font=dict(family="Syne", size=13, color="#7eb8e8")),
    xaxis=dict(**AXIS_STYLE),
    yaxis=dict(title="Enterprises", **AXIS_STYLE),
    margin=dict(l=10,r=60,t=40,b=10), height=320,
    legend=dict(font=dict(color="#8ab4d8", size=11), bgcolor="rgba(0,0,0,0)", orientation="h", y=1.08),
    barmode="group",
)
st.plotly_chart(fig9, use_container_width=True)
xl_btn(year_df.merge(emp_year, on="Year", how="left").rename(columns={"Count":"Enterprises","TotalEmp":"Employees"}) if emp_year is not None else year_df.rename(columns={"Count":"Enterprises"}), "year_trend")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── ROW 6: NIC × District heatmap ────────────────────────────────────────────
dist_col = "Dic_Name" if "Dic_Name" in bool_fdf.columns else "DISTRICT_NAME"
if dist_col in bool_fdf.columns:
    st.markdown('<div class="section-head">NIC × District Heatmap</div>', unsafe_allow_html=True)
    top_nic_ids = nic_count["NIC5DigitId"].head(15).tolist()
    top_dists   = (
        bool_fdf.groupby(dist_col)["UdyogAadharNo"].nunique()
        .sort_values(ascending=False).head(15).index.tolist()
    )
    heat_df = (
        bool_fdf[bool_fdf["NIC5DigitId"].isin(top_nic_ids) & bool_fdf[dist_col].isin(top_dists)]
        .groupby(["NIC5DigitId", dist_col])["UdyogAadharNo"].nunique()
        .reset_index().rename(columns={"UdyogAadharNo":"Count"})
    )
    if not heat_df.empty:
        pivot = heat_df.pivot(index="NIC5DigitId", columns=dist_col, values="Count").fillna(0)
        fig10 = go.Figure(go.Heatmap(
            z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
            colorscale=[[0,"#0a0f1e"],[0.3,"#0d2a4a"],[0.7,"#1565c0"],[1,"#4fc3f7"]],
            text=pivot.values.astype(int),
            texttemplate="%{text}",
            textfont=dict(size=10, color="#c8d8f0"),
            hoverongaps=False,
        ))
        fig10.update_layout(
            **CHART_THEME,
            title=dict(text="NIC Code × DIC District (enterprise count)", font=dict(family="Syne", size=13, color="#7eb8e8")),
            xaxis=dict(**AXIS_STYLE, tickangle=-30),
            yaxis=dict(**AXIS_STYLE),
            margin=dict(l=10,r=10,t=40,b=60),
            height=420,
        )
        st.plotly_chart(fig10, use_container_width=True)
        xl_btn(heat_df.rename(columns={"Count":"Enterprises"}), "nic_district_heatmap")

st.markdown('<div class="fancy-divider"></div>', unsafe_allow_html=True)

# ── Full NIC Table + Download ─────────────────────────────────────────────────
st.markdown('<div class="section-head">Full NIC Code Summary Table</div>', unsafe_allow_html=True)

# Build rich table
base = bool_fdf.groupby(["NIC5DigitId","NICDescription"]).agg(
    EnterpriseCount=("UdyogAadharNo","nunique"),
).reset_index()

if "TotalEmp" in bool_fdf.columns:
    emp_nic = (
        bool_fdf.drop_duplicates(["UdyogAadharNo","NIC5DigitId"])
        .groupby("NIC5DigitId")["TotalEmp"].sum().reset_index()
        .rename(columns={"TotalEmp":"TotalEmployees"})
    )
    base = base.merge(emp_nic, on="NIC5DigitId", how="left")

if "InvestmentCost" in bool_fdf.columns:
    inv_nic = (
        bool_fdf.drop_duplicates(["UdyogAadharNo","NIC5DigitId"])
        .groupby("NIC5DigitId")["InvestmentCost"].sum().reset_index()
    )
    base = base.merge(inv_nic, on="NIC5DigitId", how="left")
    base["InvestmentCost"] = base["InvestmentCost"].apply(lambda v: fmt_num(v) if pd.notna(v) else "—")

if "NetTurnover" in bool_fdf.columns:
    tur_nic = (
        bool_fdf.drop_duplicates(["UdyogAadharNo","NIC5DigitId"])
        .groupby("NIC5DigitId")["NetTurnover"].sum().reset_index()
    )
    base = base.merge(tur_nic, on="NIC5DigitId", how="left")
    base["NetTurnover"] = base["NetTurnover"].apply(lambda v: fmt_num(v) if pd.notna(v) else "—")

base = base.sort_values("EnterpriseCount", ascending=False).reset_index(drop=True)
st.dataframe(base, use_container_width=True, hide_index=True, height=320)

xl_btn(base, "nic_full_summary", "⬇️  Download NIC Summary Excel")
