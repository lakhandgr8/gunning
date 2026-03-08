import streamlit as st
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader
import pandas as pd
from datetime import datetime, date
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import os
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# CONFIGURATION & CONSTANTS
# ============================================================
PAGE_TITLE  = "Gunning Mass Stock Register"
PAGE_ICON   = "🏭"
DATE_FORMAT = '%d-%b-%Y'
DATA_FILE   = "Google Sheets (StockData tab)"
DEFAULT_LOW_STOCK_THRESHOLD = 500.0

REPAIR_ZONES = [
    "Slagdoor", "E1 Hotspot", "E2 Hotspot",
    "E3 Hotspot", "Elbow Area", "Multiple Zones", "Other"
]

STOCK_COLUMNS = [
    'Entry ID', 'Date', 'Entry Type',
    'Opening Stock (Kg)',
    'Received from Store (MT)', 'Received from Store (Kg)',
    'Used for Sidewall Repair (Kg)',
    'Closing Stock (Kg)',
    'Total Consumption Till Date (Kg)',
    'Remarks'
]

# ============================================================
# PAGE CONFIG (Must be first Streamlit command)
# ============================================================
st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON, layout="wide")


# ============================================================
# AUTHENTICATION LAYER
# ============================================================
try:
    with open('config.yaml') as file:
        config = yaml.load(file, Loader=SafeLoader)
except FileNotFoundError:
    st.error("🚨 Missing 'config.yaml' file. Please create it to enable login.")
    st.stop()

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

try:
    authenticator.login()
except Exception as e:
    st.error(e)

if st.session_state["authentication_status"] is False:
    st.error('Username/password is incorrect')
    st.stop()
elif st.session_state["authentication_status"] is None:
    st.warning('Please enter your username and password to access the register.')
    st.stop()


# ============================================================
# CSS
# ============================================================
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background:#f5f7fa; }
[data-testid="stSidebar"]          { background:#1a1a2e; }
[data-testid="stSidebar"] * { color:#e0e0e0 !important; }
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3       { color:#ffffff !important; }

.main-header {
    background: linear-gradient(135deg,#1a1a2e 0%,#16213e 50%,#0f3460 100%);
    padding:28px 20px; border-radius:14px;
    margin-bottom:24px; text-align:center; color:white;
    box-shadow:0 4px 20px rgba(0,0,0,.3);
}
.main-header h1 { font-size:2rem; margin:0 0 6px 0; letter-spacing:1px; }
.main-header p  { margin:0; opacity:.85; font-size:1rem; }

.kpi-card {
    background:white; border-radius:12px;
    padding:18px 16px; text-align:center;
    box-shadow:0 2px 10px rgba(0,0,0,.08);
    border-top:4px solid #0f3460; margin-bottom:8px;
}
.kpi-card .kpi-val   { font-size:1.6rem; font-weight:700; color:#0f3460; }
.kpi-card .kpi-label { font-size:.78rem; color:#777;
                        text-transform:uppercase; letter-spacing:.5px; margin-top:4px; }

.alert-critical {
    background:#ffe0e0; border-left:5px solid #e53935;
    padding:12px 16px; border-radius:8px;
    margin-bottom:12px; font-weight:600; color:#b71c1c;
}
.alert-warning {
    background:#fff8e1; border-left:5px solid #ffb300;
    padding:12px 16px; border-radius:8px;
    margin-bottom:12px; font-weight:600; color:#e65100;
}
.alert-success {
    background:#e8f5e9; border-left:5px solid #43a047;
    padding:12px 16px; border-radius:8px;
    margin-bottom:12px; font-weight:600; color:#1b5e20;
}

.download-card {
    background:white; border-radius:12px;
    padding:24px; margin-bottom:16px;
    box-shadow:0 2px 10px rgba(0,0,0,.08);
    border-left:5px solid #0f3460;
}
.download-card h3 { color:#0f3460; margin-bottom:6px; }
.download-card p  { color:#666; font-size:13px; margin-bottom:14px; }

.report-section {
    background:#f8f9ff; border-radius:12px;
    padding:20px; margin-bottom:16px;
    border:1px solid #e0e4ff;
}

.badge {
    display:inline-block; border-radius:12px;
    padding:2px 9px; font-size:11px; font-weight:700;
}
.badge-receipt     { background:#e8f5e9; color:#2e7d32; border:1px solid #a5d6a7; }
.badge-consumption { background:#fff3e0; color:#e65100; border:1px solid #ffcc80; }
.badge-initial     { background:#e3f2fd; color:#1565c0; border:1px solid #90caf9; }

.del-btn button {
    background:#ef5350 !important; color:white !important;
    border:none !important; border-radius:6px !important;
    font-size:14px !important; min-height:32px !important;
    width:100% !important; cursor:pointer !important;
}
.del-btn button:hover { background:#c62828 !important; }

.bulk-bar {
    background:#e8eaf6; border:1.5px solid #3f51b5;
    border-radius:10px; padding:10px 16px; margin:8px 0;
}
.confirm-box {
    background:#fff5f5; border:2px solid #ef5350;
    border-radius:10px; padding:14px 18px;
    margin:8px 0; color:#b71c1c; font-weight:600;
}
.form-section {
    background:white; border-radius:12px;
    padding:20px; margin-bottom:16px;
    box-shadow:0 2px 8px rgba(0,0,0,.07);
}
.section-header {
    font-size:1.1rem; font-weight:700; color:#0f3460;
    border-bottom:2px solid #e0e0e0;
    padding-bottom:6px; margin-bottom:14px;
}
.unit-note {
    font-size:12px; color:#888;
    margin-top:-8px; margin-bottom:10px;
}
.filename-tip {
    background:#e3f2fd; border-radius:8px;
    padding:10px 14px; font-size:12px;
    color:#1565c0; margin-bottom:10px;
}
div[data-testid="stHorizontalBlock"] { align-items:center; }
</style>
""", unsafe_allow_html=True)


# ============================================================
# DATA PERSISTENCE — Direct gspread (no st-gsheets-connection)
# ============================================================

GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _get_gspread_client():
    """
    Build an authenticated gspread client each time — no caching,
    avoids stale-credential bugs with st.cache_resource.
    """
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    # Strip keys that are not part of the service-account JSON
    for key in ("spreadsheet", "type", "allow_programmatic_writes"):
        creds_dict.pop(key, None)
    creds = Credentials.from_service_account_info(creds_dict, scopes=GSHEETS_SCOPES)
    return gspread.authorize(creds)


def _get_worksheet():
    """Return the StockData worksheet object."""
    client          = _get_gspread_client()
    spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
    sh              = client.open_by_url(spreadsheet_url)
    try:
        ws = sh.worksheet("StockData")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="StockData", rows=1000, cols=20)
        ws.append_row(STOCK_COLUMNS)
    return ws


def _sanitize_for_sheets(df: pd.DataFrame) -> pd.DataFrame:
    """Replace NaN/None/NaT with safe defaults before writing."""
    out = df.copy()
    out['Date'] = pd.to_datetime(out['Date']).dt.strftime('%Y-%m-%d %H:%M:%S')
    num_cols = [
        'Entry ID', 'Opening Stock (Kg)',
        'Received from Store (MT)', 'Received from Store (Kg)',
        'Used for Sidewall Repair (Kg)', 'Closing Stock (Kg)',
        'Total Consumption Till Date (Kg)',
    ]
    for col in num_cols:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors='coerce').fillna(0.0)
    for col in ['Entry Type', 'Remarks']:
        if col in out.columns:
            out[col] = out[col].fillna('').astype(str)
    return out


def load_data() -> pd.DataFrame:
    """Read all rows from the StockData worksheet."""
    try:
        ws      = _get_worksheet()
        records = ws.get_all_records(expected_headers=STOCK_COLUMNS)

        if not records:
            return pd.DataFrame(columns=STOCK_COLUMNS)

        df = pd.DataFrame(records)
        df = df.dropna(how='all').reset_index(drop=True)

        if len(df) == 0:
            return pd.DataFrame(columns=STOCK_COLUMNS)

        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

        num_cols = [
            'Entry ID', 'Opening Stock (Kg)',
            'Received from Store (MT)', 'Received from Store (Kg)',
            'Used for Sidewall Repair (Kg)', 'Closing Stock (Kg)',
            'Total Consumption Till Date (Kg)',
        ]
        for col in num_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

        df['Entry ID'] = df['Entry ID'].astype(int)
        df['Remarks']  = df['Remarks'].fillna('').astype(str)
        return df

    except Exception as e:
        import traceback
        st.error(f"❌ LOAD ERROR — {type(e).__name__}: {e}")
        st.code(traceback.format_exc())
        return pd.DataFrame(columns=STOCK_COLUMNS)


def save_data(df: pd.DataFrame) -> bool:
    """
    Overwrite the StockData worksheet with the full DataFrame.
    Uses clear() + update() for an atomic write.
    """
    try:
        if df is None or len(df) == 0:
            df = pd.DataFrame(columns=STOCK_COLUMNS)

        df_out = _sanitize_for_sheets(df)
        ws     = _get_worksheet()

        header = STOCK_COLUMNS
        rows   = df_out[STOCK_COLUMNS].values.tolist()

        safe_rows = []
        for row in rows:
            safe_row = []
            for val in row:
                if val is None or (isinstance(val, float) and pd.isna(val)):
                    safe_row.append('')
                elif isinstance(val, (int, float, str, bool)):
                    safe_row.append(val)
                else:
                    safe_row.append(str(val))
            safe_rows.append(safe_row)

        ws.clear()
        ws.update('A1', [header] + safe_rows)
        return True

    except Exception as e:
        import traceback
        st.error(f"❌ SAVE ERROR — {type(e).__name__}: {e}")
        st.code(traceback.format_exc())
        return False


# ============================================================
# HELPERS
# ============================================================
def get_current_stock() -> float:
    if len(st.session_state.stock_data) > 0:
        return float(st.session_state.stock_data.iloc[-1]['Closing Stock (Kg)'])
    return 0.0


def get_total_consumption() -> float:
    if len(st.session_state.stock_data) > 0:
        return float(
            st.session_state.stock_data.iloc[-1]['Total Consumption Till Date (Kg)']
        )
    return 0.0


def get_next_id() -> int:
    if len(st.session_state.stock_data) > 0:
        return int(st.session_state.stock_data['Entry ID'].max()) + 1
    return 1


def is_init() -> bool:
    return st.session_state.get('initial_stock_set', False)


def stock_status(stock: float, thr: float) -> str:
    if stock <= 0:   return "critical"
    if stock < thr:  return "warning"
    return "good"


def fmt_date(dt) -> str:
    if pd.isna(dt):
        return ''
    return pd.to_datetime(dt).strftime(DATE_FORMAT)


def mt2kg(mt: float) -> float:
    return mt * 1000.0


def kg2mt(kg: float) -> float:
    return kg / 1000.0


def valid_consumption_multiples(max_kg: float) -> list:
    top = int(max_kg // 100) * 100
    if top < 100:
        return []
    return list(range(100, top + 1, 100))


def kpi(label: str, value: str, sub: str = "") -> str:
    sub_html = (
        "<div style='font-size:11px;color:#aaa;margin-top:3px'>"
        + sub + "</div>" if sub else ""
    )
    return (
        "<div class='kpi-card'>"
        "<div class='kpi-val'>" + value + "</div>"
        "<div class='kpi-label'>" + label + "</div>"
        + sub_html + "</div>"
    )


def _border_color(entry_type: str) -> str:
    mapping = {
        'Receipt':     '#43a047',
        'Consumption': '#fb8c00',
        'Initial':     '#1e88e5',
    }
    return mapping.get(entry_type, '#cccccc')


def _badge(entry_type: str) -> str:
    mapping = {
        'Receipt':     'badge-receipt',
        'Consumption': 'badge-consumption',
        'Initial':     'badge-initial',
    }
    cls = mapping.get(entry_type, '')
    return "<span class='badge " + cls + "'>" + entry_type + "</span>"


# ============================================================
# REPORT GENERATION HELPERS
# ============================================================
def prepare_export_df(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out['Date'] = out['Date'].apply(fmt_date)
    for col in ['Opening Stock (Kg)', 'Received from Store (Kg)',
                'Used for Sidewall Repair (Kg)', 'Closing Stock (Kg)',
                'Total Consumption Till Date (Kg)']:
        if col in out.columns:
            out[col] = out[col].round(2)
    return out


def build_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8')


def build_excel(df_all: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        edf = prepare_export_df(df_all)
        edf.to_excel(writer, index=False, sheet_name='All Entries')

        rcv = edf[edf['Entry Type'].isin(['Receipt', 'Initial'])]
        rcv.to_excel(writer, index=False, sheet_name='Receipts')

        con = edf[edf['Entry Type'] == 'Consumption']
        con.to_excel(writer, index=False, sheet_name='Consumption')

        cur      = get_current_stock()
        tot_rcv  = df_all[
            df_all['Entry Type'].isin(['Receipt', 'Initial'])
        ]['Received from Store (Kg)'].sum()
        tot_con  = df_all['Used for Sidewall Repair (Kg)'].sum()
        con_only = df_all[
            df_all['Entry Type'] == 'Consumption'
        ]['Used for Sidewall Repair (Kg)']
        avg_use  = con_only.mean() if len(con_only) > 0 else 0
        max_use  = con_only.max()  if len(con_only) > 0 else 0

        summary = pd.DataFrame({
            'Metric': [
                'Report Generated',
                'Total Entries',
                'Current Stock (Kg)',
                'Current Stock (MT)',
                'Total Received (Kg)',
                'Total Received (MT)',
                'Total Consumed (Kg)',
                'Avg Consumption per Entry (Kg)',
                'Max Single Consumption (Kg)',
            ],
            'Value': [
                datetime.now().strftime('%d-%b-%Y %H:%M'),
                len(df_all),
                round(cur, 2),
                round(kg2mt(cur), 3),
                round(tot_rcv, 2),
                round(kg2mt(tot_rcv), 3),
                round(tot_con, 2),
                round(avg_use, 2),
                round(max_use, 2),
            ]
        })
        summary.to_excel(writer, index=False, sheet_name='Summary')

    return buf.getvalue()


def build_consumption_report(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    con_df = df[df['Entry Type'] == 'Consumption'].copy()

    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        edf = prepare_export_df(con_df)
        edf.to_excel(writer, index=False, sheet_name='Consumption Detail')

        zone_rows = []
        for z in REPAIR_ZONES:
            mask = con_df['Remarks'].str.contains(z, case=False, na=False)
            kg   = con_df.loc[mask, 'Used for Sidewall Repair (Kg)'].sum()
            cnt  = mask.sum()
            if kg > 0:
                zone_rows.append({
                    'Repair Zone': z,
                    'No. of Entries': int(cnt),
                    'Total Used (Kg)': round(kg, 2),
                    'Total Used (MT)': round(kg2mt(kg), 3),
                })
        if zone_rows:
            zdf = pd.DataFrame(zone_rows)
            zdf.to_excel(writer, index=False, sheet_name='Zone Breakdown')

        if len(con_df) > 0:
            con_df['Month'] = pd.to_datetime(con_df['Date']).dt.to_period('M').astype(str)
            monthly = (
                con_df.groupby('Month')['Used for Sidewall Repair (Kg)']
                .agg(['sum', 'count', 'mean', 'max'])
                .reset_index()
            )
            monthly.columns = [
                'Month', 'Total Used (Kg)',
                'No. of Entries', 'Avg per Entry (Kg)', 'Max Single (Kg)'
            ]
            monthly = monthly.round(2)
            monthly.to_excel(writer, index=False, sheet_name='Monthly Summary')

    return buf.getvalue()


def build_receipt_report(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    rcv_df = df[df['Entry Type'].isin(['Receipt', 'Initial'])].copy()

    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        edf = prepare_export_df(rcv_df)
        edf.to_excel(writer, index=False, sheet_name='Receipt Detail')

        if len(rcv_df) > 0:
            rcv_df['Month'] = pd.to_datetime(rcv_df['Date']).dt.to_period('M').astype(str)
            monthly = (
                rcv_df.groupby('Month')['Received from Store (Kg)']
                .agg(['sum', 'count'])
                .reset_index()
            )
            monthly.columns = ['Month', 'Total Received (Kg)', 'No. of Receipts']
            monthly['Total Received (MT)'] = (monthly['Total Received (Kg)'] / 1000).round(3)
            monthly.to_excel(writer, index=False, sheet_name='Monthly Summary')

    return buf.getvalue()


# ============================================================
# DOWNLOAD WIDGET BUILDER
# ============================================================
def download_widget(label, data, suggested_name, mime, file_ext, help_text=""):
    clean_default = suggested_name.replace(' ', '_')
    col_name, col_btn = st.columns([3, 1])
    with col_name:
        user_name = st.text_input(
            "📝 Save as filename:",
            value=clean_default,
            key="fname_" + label.replace(" ", "_") + file_ext,
            help=help_text or "Type your preferred filename."
        )
        st.markdown(
            "<div class='filename-tip'>"
            "💡 Your browser will prompt you to choose the save location "
            "when you click the download button."
            "</div>",
            unsafe_allow_html=True
        )

    final_name = user_name.strip() or clean_default
    if not final_name.endswith(file_ext):
        final_name += file_ext

    with col_btn:
        st.markdown("<div style='margin-top:28px;'>", unsafe_allow_html=True)
        st.download_button(
            label=label, data=data,
            file_name=final_name, mime=mime,
            width='stretch'
        )
        st.markdown("</div>", unsafe_allow_html=True)

    return final_name


# ============================================================
# BUSINESS LOGIC
# ============================================================
def recalculate_all(df: pd.DataFrame) -> pd.DataFrame:
    if len(df) == 0:
        return df
    df = df.sort_values('Date').reset_index(drop=True)
    cumulative = 0.0
    for i in range(len(df)):
        if i == 0:
            opening = float(df.loc[i, 'Opening Stock (Kg)'])
        else:
            opening = float(df.loc[i - 1, 'Closing Stock (Kg)'])
        df.loc[i, 'Opening Stock (Kg)']  = opening
        received = float(df.loc[i, 'Received from Store (Kg)'])
        used     = float(df.loc[i, 'Used for Sidewall Repair (Kg)'])
        df.loc[i, 'Closing Stock (Kg)']  = round(opening + received - used, 4)
        cumulative += used
        df.loc[i, 'Total Consumption Till Date (Kg)'] = round(cumulative, 4)
    return df


def do_delete(ids: list) -> tuple:
    df        = st.session_state.stock_data.copy()
    remaining = df[~df['Entry ID'].isin(ids)].reset_index(drop=True)
    noun      = 'entry' if len(ids) == 1 else 'entries'

    if len(remaining) == 0:
        st.session_state.stock_data        = pd.DataFrame(columns=STOCK_COLUMNS)
        st.session_state.initial_stock_set = False
        save_data(st.session_state.stock_data)
        return True, "All " + str(len(ids)) + " " + noun + " deleted. Register reset."

    remaining = recalculate_all(remaining)
    if (remaining['Closing Stock (Kg)'] < 0).any():
        return False, "Deletion would cause negative stock. Operation cancelled."

    st.session_state.stock_data = remaining
    save_data(remaining)
    return True, str(len(ids)) + " " + noun + " deleted and register recalculated."


def make_receipt_row(entry_date, mt: float, remarks: str) -> pd.DataFrame:
    kg  = mt2kg(mt)
    cur = get_current_stock()
    return pd.DataFrame([{
        'Entry ID':                         get_next_id(),
        'Date':                             pd.to_datetime(entry_date),
        'Entry Type':                       'Receipt',
        'Opening Stock (Kg)':               cur,
        'Received from Store (MT)':         mt,
        'Received from Store (Kg)':         kg,
        'Used for Sidewall Repair (Kg)':    0.0,
        'Closing Stock (Kg)':               round(cur + kg, 4),
        'Total Consumption Till Date (Kg)': get_total_consumption(),
        'Remarks':                          remarks or ''
    }])


def make_consumption_row(entry_date, used_kg: float, remarks: str) -> pd.DataFrame:
    cur = get_current_stock()
    return pd.DataFrame([{
        'Entry ID':                         get_next_id(),
        'Date':                             pd.to_datetime(entry_date),
        'Entry Type':                       'Consumption',
        'Opening Stock (Kg)':               cur,
        'Received from Store (MT)':         0.0,
        'Received from Store (Kg)':         0.0,
        'Used for Sidewall Repair (Kg)':    used_kg,
        'Closing Stock (Kg)':               round(cur - used_kg, 4),
        'Total Consumption Till Date (Kg)': round(get_total_consumption() + used_kg, 4),
        'Remarks':                          remarks or ''
    }])


def append_row(row: pd.DataFrame):
    st.session_state.stock_data = pd.concat(
        [st.session_state.stock_data, row], ignore_index=True
    )
    save_data(st.session_state.stock_data)


# ============================================================
# SESSION STATE
# ============================================================
if 'stock_data' not in st.session_state:
    loaded = load_data()
    st.session_state.stock_data = loaded
if 'initial_stock_set' not in st.session_state:
    st.session_state.initial_stock_set = len(st.session_state.stock_data) > 0
if 'low_thr' not in st.session_state:
    st.session_state.low_thr = DEFAULT_LOW_STOCK_THRESHOLD
if 'selected_ids' not in st.session_state:
    st.session_state.selected_ids = set()
if 'pending_del' not in st.session_state:
    st.session_state.pending_del = []


# ============================================================
# HEADER
# ============================================================
st.markdown(
    "<div class='main-header'>"
    "<h1>🏭 GUNNING MASS STOCK REGISTER</h1>"
    "<p>EAF Sidewall Repair — Stock Management System</p>"
    "</div>",
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 📋 Navigation")
    action = st.radio("Navigation", [
        "🏠 Dashboard",
        "📥 Receive Stock",
        "🔥 Log Consumption",
        "📖 View Register",
        "✏️ Edit Entry",
        "📊 Analytics",
        "📋 Reports",
        "💾 Download Data",
        "📤 Import Data",
    ], label_visibility="collapsed")

    st.markdown("---")
    authenticator.logout('Log Out', 'sidebar')
    st.markdown(f"**Logged in as:** {st.session_state['name']}")

    st.markdown("---")
    st.markdown("### ⚙️ Settings")
    st.session_state.low_thr = st.number_input(
        "Low Stock Threshold (Kg):",
        min_value=0.0,
        value=float(st.session_state.low_thr),
        step=50.0
    )

    st.markdown("---")
    st.markdown("### 🔧 Connection Debug")
    if st.button("🔍 Test Google Sheets", key="debug_btn"):
        with st.spinner("Testing..."):
            try:
                # Step 1: secrets present?
                gsheet_secrets = dict(st.secrets["connections"]["gsheets"])
                st.success("✅ Secrets found")

                # Step 2: required keys?
                required = ["spreadsheet", "private_key", "client_email"]
                missing  = [k for k in required if k not in gsheet_secrets]
                if missing:
                    st.error(f"❌ Missing secret keys: {missing}")
                else:
                    st.success("✅ All required keys present")
                    st.caption(f"client_email: {gsheet_secrets.get('client_email','?')}")
                    st.caption(f"spreadsheet: {gsheet_secrets.get('spreadsheet','?')[:60]}...")

                # Step 3: can we authenticate?
                from google.oauth2.service_account import Credentials
                import gspread as _gs
                creds_dict = {k: v for k, v in gsheet_secrets.items()
                              if k not in ("spreadsheet","type","allow_programmatic_writes")}
                scopes = ["https://www.googleapis.com/auth/spreadsheets",
                          "https://www.googleapis.com/auth/drive"]
                creds  = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                client = _gs.authorize(creds)
                st.success("✅ Google auth OK")

                # Step 4: can we open the sheet?
                sh = client.open_by_url(gsheet_secrets["spreadsheet"])
                st.success(f"✅ Opened spreadsheet: '{sh.title}'")

                # Step 5: can we find/read the worksheet?
                try:
                    ws = sh.worksheet("StockData")
                    rows = ws.get_all_records()
                    st.success(f"✅ StockData worksheet found — {len(rows)} data rows")
                except Exception as e:
                    st.warning(f"⚠️ StockData tab not found: {e}")

                # Step 6: test write
                try:
                    ws = sh.worksheet("StockData")
                    existing = ws.acell('A1').value
                    st.success(f"✅ Read A1: '{existing}'")
                    st.info("✅ All checks passed — connection is working!")
                except Exception as e:
                    st.error(f"❌ Read test failed: {e}")

            except KeyError as e:
                st.error(f"❌ secrets.toml missing key: {e}")
            except Exception as e:
                st.error(f"❌ Connection failed: {type(e).__name__}: {e}")

    st.markdown("---")
    st.markdown("### 📊 Quick Status")
    if len(st.session_state.stock_data) > 0:
        _c = get_current_stock()
        _s = stock_status(_c, st.session_state.low_thr)
        if   _s == "critical": st.error(  "🚨 CRITICAL: " + f"{_c:.0f}" + " Kg")
        elif _s == "warning":  st.warning("⚠️ LOW: "      + f"{_c:.0f}" + " Kg")
        else:                  st.success("✅ OK: "        + f"{_c:.0f}" + " Kg")
        st.info(
            "**Entries:** " + str(len(st.session_state.stock_data)) + "\n\n"
            "**Stock:** " + f"{_c:.2f}" + " Kg  |  " + f"{kg2mt(_c):.3f}" + " MT\n\n"
            "**Consumed:** " + f"{get_total_consumption():.2f}" + " Kg"
        )
    else:
        st.info("No data yet.")


# ============================================================
# MODULE 1 — INITIAL SETUP
# ============================================================
def render_initial_setup():
    st.header("🚀 Initial Stock Setup")
    st.info("No data found. Set the opening stock to begin tracking.")

    with st.form("init_form"):
        c1, c2 = st.columns(2)
        with c1:
            init_mt = st.number_input(
                "Opening Stock (MT):",
                min_value=0.0, value=0.0, step=1.0,
                help="1 MT = 1 000 Kg"
            )
            st.markdown(
                "<div class='unit-note'>= "
                + str(int(mt2kg(init_mt))) + " Kg</div>",
                unsafe_allow_html=True
            )
        with c2:
            init_date = st.date_input("Start Date:", value=date.today())

        remarks = st.text_input(
            "Remarks:",
            placeholder="e.g. Carry forward from previous period"
        )
        ok = st.form_submit_button(
            "✅ Set Initial Stock", type="primary", width='stretch'
        )

    if ok:
        kg = mt2kg(init_mt)
        row = pd.DataFrame([{
            'Entry ID':                         1,
            'Date':                             pd.to_datetime(init_date),
            'Entry Type':                       'Initial',
            'Opening Stock (Kg)':               0.0,
            'Received from Store (MT)':         init_mt,
            'Received from Store (Kg)':         kg,
            'Used for Sidewall Repair (Kg)':    0.0,
            'Closing Stock (Kg)':               kg,
            'Total Consumption Till Date (Kg)': 0.0,
            'Remarks':                          remarks or 'Initial Stock Setup'
        }])
        st.session_state.stock_data        = row
        st.session_state.initial_stock_set = True
        save_data(row)
        st.success(
            "✅ Initial stock set: **"
            + str(int(init_mt)) + " MT ("
            + str(int(kg)) + " Kg)**"
        )
        st.balloons()
        st.rerun()


# ============================================================
# MODULE 2 — DASHBOARD
# ============================================================
def render_dashboard():
    if len(st.session_state.stock_data) == 0:
        render_initial_setup()
        return

    df         = st.session_state.stock_data.copy()
    df['Date'] = pd.to_datetime(df['Date'])
    rcv_df     = df[df['Entry Type'].isin(['Receipt', 'Initial'])]
    con_df     = df[df['Entry Type'] == 'Consumption']

    cur      = get_current_stock()
    tot_con  = get_total_consumption()
    tot_rcv  = rcv_df['Received from Store (Kg)'].sum()
    avg_use  = (
        con_df['Used for Sidewall Repair (Kg)'].mean()
        if len(con_df) > 0 else 0
    )
    days_rem = (cur / avg_use) if avg_use > 0 else float('inf')
    eff      = (tot_con / tot_rcv * 100) if tot_rcv > 0 else 0
    status   = stock_status(cur, st.session_state.low_thr)

    if status == "critical":
        st.markdown(
            "<div class='alert-critical'>"
            "🚨 CRITICAL: Stock is at zero or negative!</div>",
            unsafe_allow_html=True
        )
    elif status == "warning":
        st.markdown(
            "<div class='alert-warning'>⚠️ LOW STOCK: "
            + f"{cur:.2f}" + " Kg remaining (threshold "
            + f"{st.session_state.low_thr:.0f}" + " Kg)</div>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            "<div class='alert-success'>✅ Stock level is healthy: "
            + f"{cur:.2f}" + " Kg</div>",
            unsafe_allow_html=True
        )

    days_label = f"{days_rem:.1f}d" if days_rem != float('inf') else "∞"
    cols = st.columns(6)
    kpi_data = [
        ("Current Stock",  f"{cur:.0f} Kg",     f"{kg2mt(cur):.3f} MT"),
        ("Total Received", f"{tot_rcv:.0f} Kg",  f"{kg2mt(tot_rcv):.2f} MT"),
        ("Total Consumed", f"{tot_con:.0f} Kg",  "cumulative"),
        ("Avg per Entry",  f"{avg_use:.0f} Kg",  "per consumption"),
        ("Stock Duration", days_label,             "at avg usage"),
        ("Utilization",    f"{eff:.1f}%",          "consumed / received"),
    ]
    for col, (lbl, val, sub) in zip(cols, kpi_data):
        col.markdown(kpi(lbl, val, sub), unsafe_allow_html=True)

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(
            "<div class='section-header'>📥 Recent Receipts</div>",
            unsafe_allow_html=True
        )
        if len(rcv_df) > 0:
            rd = rcv_df.tail(5).copy()
            rd['Date'] = rd['Date'].apply(fmt_date)
            st.dataframe(
                rd[['Date', 'Received from Store (MT)',
                    'Received from Store (Kg)',
                    'Closing Stock (Kg)', 'Remarks']],
                width='stretch', hide_index=True
            )
        else:
            st.info("No receipts yet.")

    with c2:
        st.markdown(
            "<div class='section-header'>🔥 Recent Consumption</div>",
            unsafe_allow_html=True
        )
        if len(con_df) > 0:
            cd = con_df.tail(5).copy()
            cd['Date'] = cd['Date'].apply(fmt_date)
            st.dataframe(
                cd[['Date', 'Used for Sidewall Repair (Kg)',
                    'Closing Stock (Kg)',
                    'Total Consumption Till Date (Kg)', 'Remarks']],
                width='stretch', hide_index=True
            )
        else:
            st.info("No consumption yet.")

    if len(df) >= 2:
        st.markdown("---")
        st.markdown(
            "<div class='section-header'>📈 Stock Level Trend</div>",
            unsafe_allow_html=True
        )
        marker_colors  = [
            '#43a047' if t in ['Receipt', 'Initial'] else '#fb8c00'
            for t in df['Entry Type']
        ]
        marker_symbols = [
            'triangle-up' if t in ['Receipt', 'Initial'] else 'triangle-down'
            for t in df['Entry Type']
        ]
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df['Date'], y=df['Closing Stock (Kg)'],
            mode='lines+markers', name='Stock',
            line=dict(color='#0f3460', width=2.5),
            fill='tozeroy', fillcolor='rgba(15,52,96,.08)',
            marker=dict(size=7, color=marker_colors, symbol=marker_symbols)
        ))
        fig.add_hline(
            y=st.session_state.low_thr,
            line_dash="dash", line_color="#e53935",
            annotation_text="Threshold: "
            + f"{st.session_state.low_thr:.0f}" + " Kg"
        )
        fig.update_layout(
            height=300, margin=dict(l=0, r=0, t=10, b=0),
            xaxis_title="Date", yaxis_title="Stock (Kg)",
            hovermode='x unified',
            plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# MODULE 3 — RECEIVE STOCK
# ============================================================
def render_receive():
    st.header("📥 Receive Stock from Store")
    if not is_init():
        render_initial_setup()
        return

    cur = get_current_stock()
    c1, c2 = st.columns(2)
    c1.markdown(kpi("Current Stock (Kg)", f"{cur:.2f}"), unsafe_allow_html=True)
    c2.markdown(kpi("Current Stock (MT)", f"{kg2mt(cur):.3f}"), unsafe_allow_html=True)
    st.markdown("---")

    with st.form("recv_form", clear_on_submit=True):
        st.markdown(
            "<div class='section-header'>📋 Receipt Details</div>",
            unsafe_allow_html=True
        )
        col1, col2 = st.columns(2)
        with col1:
            e_date  = st.date_input("📅 Date of Receipt:", value=date.today())
            raw_mt  = st.number_input(
                "📦 Quantity Received (MT):",
                min_value=1.0, value=1.0, step=1.0,
                help="Receipt must be in whole Metric Tonnes (1 MT = 1 000 Kg)"
            )
            recv_mt = float(int(raw_mt))
            recv_kg = mt2kg(recv_mt)
            st.markdown(
                "<div class='unit-note'>✅ "
                + str(int(recv_mt)) + " MT = <b>"
                + str(int(recv_kg)) + " Kg</b></div>",
                unsafe_allow_html=True
            )
        with col2:
            challan = st.text_input(
                "🧾 Challan / GRN No.:", placeholder="GRN-2024-001"
            )
            remarks = st.text_area("📝 Remarks:", height=88)

        st.markdown(
            "<div class='section-header'>📊 Preview</div>",
            unsafe_allow_html=True
        )
        new_close = cur + recv_kg
        p1, p2, p3, p4 = st.columns(4)
        p1.info("**Opening**\n\n" + f"{cur:.2f} Kg")
        p2.success(
            "**Receiving**\n\n+" + str(int(recv_mt))
            + " MT\n\n+" + str(int(recv_kg)) + " Kg"
        )
        p3.success("**New Closing**\n\n" + f"{new_close:.2f} Kg")
        p4.success("**= MT**\n\n" + f"{kg2mt(new_close):.3f} MT")

        sub = st.form_submit_button(
            "✅ Record Receipt", type="primary", width='stretch'
        )

    if sub:
        rmk = (remarks + " | Challan: " + challan) if challan else remarks
        append_row(make_receipt_row(e_date, recv_mt, rmk))
        st.success(
            "✅ Receipt of **" + str(int(recv_mt)) + " MT ("
            + str(int(recv_kg)) + " Kg)** recorded on "
            + fmt_date(e_date) + "!"
        )
        st.balloons()
        st.rerun()


# ============================================================
# MODULE 4 — LOG CONSUMPTION
# ============================================================
def render_consumption():
    st.header("🔥 Log Consumption")
    if not is_init():
        render_initial_setup()
        return

    cur    = get_current_stock()
    status = stock_status(cur, st.session_state.low_thr)

    if status == "critical":
        st.error("🚨 No stock available! Please receive stock first.")
        return
    elif status == "warning":
        st.markdown(
            "<div class='alert-warning'>⚠️ Low stock: "
            + f"{cur:.2f}" + " Kg remaining.</div>",
            unsafe_allow_html=True
        )

    c1, c2 = st.columns(2)
    c1.markdown(kpi("Current Stock (Kg)", f"{cur:.2f}"), unsafe_allow_html=True)
    c2.markdown(kpi("Current Stock (MT)", f"{kg2mt(cur):.3f}"), unsafe_allow_html=True)
    st.markdown("---")

    multiples = valid_consumption_multiples(cur)
    if not multiples:
        st.error(
            "❌ Insufficient stock for minimum consumption (100 Kg). "
            "Current: " + f"{cur:.2f}" + " Kg."
        )
        return

    with st.form("con_form", clear_on_submit=True):
        st.markdown(
            "<div class='section-header'>📋 Consumption Details</div>",
            unsafe_allow_html=True
        )
        col1, col2 = st.columns(2)
        with col1:
            e_date  = st.date_input("📅 Date:", value=date.today())
            used_kg = st.selectbox(
                "🔥 Quantity Used (Kg):",
                options=multiples,
                help="Must be in multiples of 100 Kg"
            )
            st.markdown(
                "<div class='unit-note'>= "
                + f"{kg2mt(used_kg):.2f}" + " MT</div>",
                unsafe_allow_html=True
            )
        with col2:
            heat_no     = st.text_input(
                "🔢 Heat Number:", placeholder="H-2024-0145"
            )
            repair_zone = st.selectbox("📍 Repair Zone:", REPAIR_ZONES)
            remarks     = st.text_area("📝 Remarks:", height=60)

        st.markdown(
            "<div class='section-header'>📊 Preview</div>",
            unsafe_allow_html=True
        )
        new_close = cur - used_kg
        new_tot   = get_total_consumption() + used_kg
        p1, p2, p3, p4 = st.columns(4)
        p1.info("**Opening**\n\n" + f"{cur:.2f} Kg")
        p2.warning(
            "**Using**\n\n-" + str(int(used_kg))
            + " Kg\n\n-" + f"{kg2mt(used_kg):.2f} MT"
        )
        if new_close < st.session_state.low_thr:
            p3.error("**New Closing**\n\n" + f"{new_close:.2f} Kg ⚠️")
        else:
            p3.success("**New Closing**\n\n" + f"{new_close:.2f} Kg")
        p4.info("**Total Consumed**\n\n" + f"{new_tot:.2f} Kg")

        sub = st.form_submit_button(
            "✅ Record Consumption", type="primary", width='stretch'
        )

    if sub:
        rmk = (remarks + " | Heat: " + heat_no
               + " | Zone: " + repair_zone)
        append_row(make_consumption_row(e_date, float(used_kg), rmk))
        st.success(
            "✅ Consumption of **" + str(int(used_kg))
            + " Kg** at **" + repair_zone
            + "** (Heat " + heat_no + ") recorded!"
        )
        st.balloons()
        st.rerun()


# ============================================================
# MODULE 5 — VIEW REGISTER
# ============================================================
def render_register_table(rows_df: pd.DataFrame, key_prefix: str):
    if len(rows_df) == 0:
        st.info("No entries to display.")
        return

    visible_ids = set(int(r['Entry ID']) for _, r in rows_df.iterrows())
    n_sel       = len(st.session_state.selected_ids)

    if n_sel > 0:
        st.markdown("<div class='bulk-bar'>", unsafe_allow_html=True)
        ba1, ba2, ba3, _ = st.columns([2.5, 1.5, 1.5, 4])
        ba1.markdown("**" + str(n_sel) + " row(s) selected**")
        with ba2:
            if st.button(
                "🗑️ Delete Selected",
                key=key_prefix + "_bulk_del",
                type="primary",
                width='stretch'
            ):
                st.session_state.pending_del = list(st.session_state.selected_ids)
                st.rerun()
        with ba3:
            if st.button(
                "✖ Clear",
                key=key_prefix + "_clear",
                width='stretch'
            ):
                st.session_state.selected_ids = set()
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.pending_del:
        n    = len(st.session_state.pending_del)
        noun = 'entry' if n == 1 else 'entries'
        st.markdown(
            "<div class='confirm-box'>⚠️ Confirm deletion of <b>"
            + str(n) + " " + noun
            + "</b>? Register will be recalculated.</div>",
            unsafe_allow_html=True
        )
        cc1, cc2, _ = st.columns([1.2, 1.2, 6])
        with cc1:
            if st.button(
                "✅ Confirm Delete",
                key=key_prefix + "_confirm",
                type="primary",
                width='stretch'
            ):
                ids = st.session_state.pending_del[:]
                ok, msg = do_delete(ids)
                st.session_state.pending_del  = []
                st.session_state.selected_ids = set()
                if ok:
                    st.success("✅ " + msg)
                else:
                    st.error("❌ " + msg)
                st.rerun()
        with cc2:
            if st.button(
                "❌ Cancel",
                key=key_prefix + "_cancel_bulk",
                width='stretch'
            ):
                st.session_state.pending_del = []
                st.rerun()

    st.markdown("")

    W = [0.30, 0.40, 0.80, 0.80, 0.85,
         0.75, 0.85, 0.80, 0.85, 0.95, 2.0, 0.38]
    H = ["☑", "ID", "Date", "Type",
         "Open(Kg)", "Rcvd(MT)", "Rcvd(Kg)",
         "Used(Kg)", "Closing(Kg)", "CumulCons(Kg)",
         "Remarks", "🗑️"]

    hcols = st.columns(W)
    all_vis_sel = (
        visible_ids.issubset(st.session_state.selected_ids)
        and len(visible_ids) > 0
    )
    sel_all = hcols[0].checkbox(
        "",
        value=all_vis_sel,
        key=key_prefix + "_selall",
        help="Select / deselect all visible rows"
    )
    if sel_all and not all_vis_sel:
        st.session_state.selected_ids |= visible_ids
        st.rerun()
    if not sel_all and all_vis_sel:
        st.session_state.selected_ids -= visible_ids
        st.rerun()

    hdr_style = (
        "font-weight:700; font-size:11px; color:#555;"
        "text-transform:uppercase; letter-spacing:.4px;"
        "border-bottom:2px solid #0f3460; padding-bottom:4px;"
    )
    for hc, lbl in zip(hcols[1:], H[1:]):
        hc.markdown(
            "<div style='" + hdr_style + "'>" + lbl + "</div>",
            unsafe_allow_html=True
        )

    for _, row in rows_df.iterrows():
        eid    = int(row['Entry ID'])
        etype  = str(row['Entry Type'])
        is_sel = eid in st.session_state.selected_ids

        bc       = _border_color(etype)
        bg_color = '#fffde7' if is_sel else 'white'
        cell_style = (
            "background:" + bg_color + ";"
            "border:1px solid #e8e8e8;"
            "border-radius:7px;"
            "padding:6px 8px;"
            "margin-bottom:2px;"
            "border-left:3px solid " + bc + ";"
        )

        if (st.session_state.pending_del == [eid]
                and eid not in st.session_state.selected_ids):
            st.markdown(
                "<div class='confirm-box'>"
                "⚠️ Delete Entry <b>#" + str(eid) + "</b> ("
                + etype + " · " + fmt_date(row['Date'])
                + ")? Register will be recalculated.</div>",
                unsafe_allow_html=True
            )
            sc1, sc2, _ = st.columns([1, 1, 7])
            with sc1:
                if st.button(
                    "✅ Yes, Delete",
                    key=key_prefix + "_sconf_" + str(eid),
                    type="primary",
                    width='stretch'
                ):
                    ok, msg = do_delete([eid])
                    st.session_state.pending_del = []
                    if ok:
                        st.success("✅ " + msg)
                    else:
                        st.error("❌ " + msg)
                    st.rerun()
            with sc2:
                if st.button(
                    "❌ Cancel",
                    key=key_prefix + "_scan_" + str(eid),
                    width='stretch'
                ):
                    st.session_state.pending_del = []
                    st.rerun()

        cols = st.columns(W)

        chk = cols[0].checkbox(
            "",
            value=is_sel,
            key=key_prefix + "_chk_" + str(eid),
            help="Select row"
        )
        if chk != is_sel:
            if chk:
                st.session_state.selected_ids.add(eid)
            else:
                st.session_state.selected_ids.discard(eid)
            st.rerun()

        def cell(ci: int, content: str):
            cols[ci].markdown(
                "<div style='" + cell_style + "'>"
                + content + "</div>",
                unsafe_allow_html=True
            )

        cell(1, "<b>#" + str(eid) + "</b>")
        cell(2, fmt_date(row['Date']))
        cols[3].markdown(
            "<div style='padding:4px 0;'>" + _badge(etype) + "</div>",
            unsafe_allow_html=True
        )

        opening = float(row['Opening Stock (Kg)'])
        cell(4, f"{opening:.2f}")

        rcv_mt = float(row.get('Received from Store (MT)', 0) or 0)
        if rcv_mt > 0:
            cell(5, "<span style='color:#2e7d32;font-weight:600'>➕ "
                 + f"{rcv_mt:.1f}" + "</span>")
        else:
            cell(5, "<span style='color:#aaa'>—</span>")

        rcv_kg = float(row.get('Received from Store (Kg)', 0) or 0)
        if rcv_kg > 0:
            cell(6, "<span style='color:#2e7d32;font-weight:600'>➕ "
                 + f"{rcv_kg:.0f}" + "</span>")
        else:
            cell(6, "<span style='color:#aaa'>—</span>")

        used = float(row['Used for Sidewall Repair (Kg)'])
        if used > 0:
            cell(7, "<span style='color:#e65100;font-weight:600'>🔥 "
                 + f"{used:.0f}" + "</span>")
        else:
            cell(7, "<span style='color:#aaa'>—</span>")

        closing = float(row['Closing Stock (Kg)'])
        cell(8, "<b>" + f"{closing:.2f}" + "</b>")

        cumul = float(row['Total Consumption Till Date (Kg)'])
        cell(9, f"{cumul:.2f}")

        rmk       = str(row['Remarks'])
        rmk_short = rmk[:60] + ('…' if len(rmk) > 60 else '')
        cell(10,
             "<span style='font-size:12px;color:#555;'>"
             + rmk_short + "</span>")

        with cols[11]:
            st.markdown("<div class='del-btn'>", unsafe_allow_html=True)
            if st.button(
                "🗑️",
                key=key_prefix + "_del_" + str(eid),
                help="Delete entry #" + str(eid),
                width='stretch'
            ):
                if st.session_state.pending_del == [eid]:
                    st.session_state.pending_del = []
                else:
                    st.session_state.pending_del = [eid]
                    st.session_state.selected_ids.discard(eid)
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    noun2 = 'entry' if len(rows_df) == 1 else 'entries'
    st.caption(
        "📄 **" + str(len(rows_df)) + "** " + noun2
        + "  |  **" + str(n_sel) + "** selected"
    )


def render_register():
    st.header("📖 Stock Register")
    if len(st.session_state.stock_data) == 0:
        st.warning("No data yet.")
        return

    st.markdown(
        "<div class='section-header'>🔍 Filters</div>",
        unsafe_allow_html=True
    )
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        types = ["All"] + sorted(
            st.session_state.stock_data['Entry Type'].unique().tolist()
        )
        ftype = st.selectbox("Entry Type:", types, key="rf_type")
    with f2:
        dmin  = pd.to_datetime(st.session_state.stock_data['Date']).min().date()
        dfrom = st.date_input("From:", value=dmin, key="rf_from")
    with f3:
        dmax  = pd.to_datetime(st.session_state.stock_data['Date']).max().date()
        dto   = st.date_input("To:",   value=dmax, key="rf_to")
    with f4:
        sort  = st.selectbox(
            "Sort:", ["Newest First", "Oldest First"], key="rf_sort"
        )

    ddf = st.session_state.stock_data.copy()
    ddf['Date'] = pd.to_datetime(ddf['Date'])
    if ftype != "All":
        ddf = ddf[ddf['Entry Type'] == ftype]
    ddf = ddf[
        (ddf['Date'].dt.date >= dfrom) &
        (ddf['Date'].dt.date <= dto)
    ]
    if sort == "Newest First":
        ddf = ddf.iloc[::-1].reset_index(drop=True)

    st.markdown("---")
    tab_all, tab_rcv, tab_con = st.tabs(
        ["📋 All Entries", "📥 Receipts Only", "🔥 Consumption Only"]
    )
    with tab_all:
        render_register_table(ddf, "all")
    with tab_rcv:
        rdf = ddf[
            ddf['Entry Type'].isin(['Receipt', 'Initial'])
        ].reset_index(drop=True)
        render_register_table(rdf, "rcv")
        if len(rdf) > 0:
            tot_r = rdf['Received from Store (Kg)'].sum()
            tot_m = rdf['Received from Store (MT)'].sum()
            st.info(
                "📦 **Total in view:** "
                + f"{tot_r:.2f}" + " Kg  ("
                + f"{tot_m:.2f}" + " MT)"
            )
    with tab_con:
        cdf = ddf[
            ddf['Entry Type'] == 'Consumption'
        ].reset_index(drop=True)
        render_register_table(cdf, "con")
        if len(cdf) > 0:
            tot_c = cdf['Used for Sidewall Repair (Kg)'].sum()
            st.info("🔥 **Total in view:** " + f"{tot_c:.2f}" + " Kg")

    st.markdown("---")
    st.markdown(
        "<div class='section-header'>📊 Summary Statistics</div>",
        unsafe_allow_html=True
    )
    df_o   = st.session_state.stock_data
    t_rcv  = df_o[
        df_o['Entry Type'].isin(['Receipt', 'Initial'])
    ]['Received from Store (Kg)'].sum()
    t_used = df_o['Used for Sidewall Repair (Kg)'].sum()
    cur    = get_current_stock()
    conly  = df_o[
        df_o['Entry Type'] == 'Consumption'
    ]['Used for Sidewall Repair (Kg)']
    avg_c  = conly.mean() if len(conly) > 0 else 0
    max_c  = conly.max()  if len(conly) > 0 else 0

    s1, s2, s3, s4, s5, s6 = st.columns(6)
    s1.markdown(kpi("Total Received",  f"{t_rcv:.0f} Kg"),       unsafe_allow_html=True)
    s2.markdown(kpi("Received (MT)",   f"{kg2mt(t_rcv):.2f} MT"), unsafe_allow_html=True)
    s3.markdown(kpi("Total Consumed",  f"{t_used:.0f} Kg"),       unsafe_allow_html=True)
    s4.markdown(kpi("Current Stock",   f"{cur:.0f} Kg"),          unsafe_allow_html=True)
    s5.markdown(kpi("Avg / Entry",     f"{avg_c:.0f} Kg"),        unsafe_allow_html=True)
    s6.markdown(kpi("Max Single",      f"{max_c:.0f} Kg"),        unsafe_allow_html=True)


# ============================================================
# MODULE 6 — EDIT ENTRY
# ============================================================
def render_edit():
    st.header("✏️ Edit Entry")
    if len(st.session_state.stock_data) == 0:
        st.warning("No data available.")
        return

    df   = st.session_state.stock_data.copy()
    opts = {}
    for _, r in df.iterrows():
        eid   = int(r['Entry ID'])
        etype = str(r['Entry Type'])
        if etype in ['Receipt', 'Initial']:
            detail = "Rcvd: " + str(int(float(r['Received from Store (MT)']))) + " MT"
        else:
            detail = "Used: " + str(int(float(r['Used for Sidewall Repair (Kg)']))) + " Kg"
        label = (
            "ID " + str(eid) + " | " + fmt_date(r['Date'])
            + " | " + etype + " | " + detail
        )
        opts[label] = eid

    sel_lbl = st.selectbox("Choose entry to edit:", list(opts.keys()))
    sel_id  = opts[sel_lbl]
    sr      = df[df['Entry ID'] == sel_id].iloc[0]

    st.markdown("---")
    d1, d2, d3 = st.columns(3)
    d1.info(
        "**ID:** " + str(sr['Entry ID']) + "\n\n"
        "**Date:** " + fmt_date(sr['Date']) + "\n\n"
        "**Type:** " + str(sr['Entry Type'])
    )
    d2.info(
        "**Opening:** " + f"{float(sr['Opening Stock (Kg)']):.2f}" + " Kg\n\n"
        "**Received:** " + f"{float(sr['Received from Store (MT)']):.2f}"
        + " MT (" + f"{float(sr['Received from Store (Kg)']):.0f}" + " Kg)\n\n"
        "**Used:** " + f"{float(sr['Used for Sidewall Repair (Kg)']):.0f}" + " Kg"
    )
    d3.info(
        "**Closing:** " + f"{float(sr['Closing Stock (Kg)']):.2f}" + " Kg\n\n"
        "**Total Consumed:** "
        + f"{float(sr['Total Consumption Till Date (Kg)']):.2f}" + " Kg\n\n"
        "**Remarks:** " + str(sr['Remarks'])
    )

    st.warning("⚠️ Saving will recalculate all subsequent entries automatically.")

    with st.form("edit_form"):
        e1, e2 = st.columns(2)
        with e1:
            new_date = st.date_input(
                "Date:", value=pd.to_datetime(sr['Date']).date()
            )
            if sr['Entry Type'] in ['Receipt', 'Initial']:
                existing_mt = float(sr['Received from Store (MT)'])
                safe_mt     = max(existing_mt, 1.0)
                raw_mt      = st.number_input(
                    "Received (MT):",
                    min_value=1.0, value=safe_mt, step=1.0,
                    help="Whole Metric Tonnes only"
                )
                new_mt     = float(int(raw_mt))
                new_rcv_kg = mt2kg(new_mt)
                new_used   = 0.0
                st.markdown(
                    "<div class='unit-note'>= "
                    + str(int(new_rcv_kg)) + " Kg</div>",
                    unsafe_allow_html=True
                )
            else:
                new_mt       = 0.0
                new_rcv_kg   = 0.0
                opening_here = float(sr['Opening Stock (Kg)'])
                cur_used     = int(float(sr['Used for Sidewall Repair (Kg)']))
                top          = int(opening_here // 100) * 100
                top          = max(top, cur_used)
                multiples    = list(range(100, top + 1, 100))
                if cur_used > 0 and cur_used not in multiples:
                    multiples = sorted(set(multiples + [cur_used]))
                if not multiples:
                    multiples = [100]
                default_idx = (
                    multiples.index(cur_used) if cur_used in multiples else 0
                )
                new_used = float(st.selectbox(
                    "Used (Kg):",
                    options=multiples,
                    index=default_idx,
                    help="Consumption in multiples of 100 Kg"
                ))
                st.markdown(
                    "<div class='unit-note'>= "
                    + f"{kg2mt(new_used):.2f}" + " MT</div>",
                    unsafe_allow_html=True
                )
        with e2:
            new_remarks = st.text_area(
                "Remarks:", value=str(sr['Remarks']), height=140
            )

        save = st.form_submit_button(
            "💾 Save Changes", type="primary", width='stretch'
        )

    if save:
        idx = st.session_state.stock_data.index[
            st.session_state.stock_data['Entry ID'] == sel_id
        ].tolist()[0]
        st.session_state.stock_data.loc[idx, 'Date']                          = pd.to_datetime(new_date)
        st.session_state.stock_data.loc[idx, 'Received from Store (MT)']      = new_mt
        st.session_state.stock_data.loc[idx, 'Received from Store (Kg)']      = new_rcv_kg
        st.session_state.stock_data.loc[idx, 'Used for Sidewall Repair (Kg)'] = new_used
        st.session_state.stock_data.loc[idx, 'Remarks']                       = new_remarks

        recalc = recalculate_all(st.session_state.stock_data)
        if (recalc['Closing Stock (Kg)'] < 0).any():
            st.error("❌ Edit causes negative stock. Changes reverted.")
            st.session_state.stock_data = load_data()
        else:
            st.session_state.stock_data = recalc
            save_data(recalc)
            st.success("✅ Entry updated and register recalculated!")
            st.rerun()


# ============================================================
# MODULE 7 — ANALYTICS
# ============================================================
def render_analytics():
    st.header("📊 Analytics & Insights")
    if len(st.session_state.stock_data) == 0:
        st.warning("No data available.")
        return

    df         = st.session_state.stock_data.copy()
    df['Date'] = pd.to_datetime(df['Date'])
    rcv_df     = df[df['Entry Type'].isin(['Receipt', 'Initial'])]
    con_df     = df[df['Entry Type'] == 'Consumption']

    cur      = get_current_stock()
    tot_rcv  = rcv_df['Received from Store (Kg)'].sum()
    tot_con  = con_df['Used for Sidewall Repair (Kg)'].sum()
    avg_use  = (
        con_df['Used for Sidewall Repair (Kg)'].mean()
        if len(con_df) > 0 else 0
    )
    days_rem   = (cur / avg_use) if avg_use > 0 else float('inf')
    eff        = (tot_con / tot_rcv * 100) if tot_rcv > 0 else 0
    days_label = f"{days_rem:.1f}d" if days_rem != float('inf') else "∞"

    k_cols = st.columns(6)
    kpi_list = [
        ("Current Stock",  f"{cur:.0f} Kg",     f"{kg2mt(cur):.3f} MT"),
        ("Total Received", f"{tot_rcv:.0f} Kg",  f"{kg2mt(tot_rcv):.2f} MT"),
        ("Total Consumed", f"{tot_con:.0f} Kg",  "cumulative"),
        ("Avg/Entry",      f"{avg_use:.0f} Kg",  "per consumption"),
        ("Stock Duration", days_label,             "at avg usage"),
        ("Utilization",    f"{eff:.1f}%",          "consumed / received"),
    ]
    for kc, (lbl, val, sub) in zip(k_cols, kpi_list):
        kc.markdown(kpi(lbl, val, sub), unsafe_allow_html=True)

    st.markdown("---")
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown(
            "<div class='section-header'>📈 Stock Level Over Time</div>",
            unsafe_allow_html=True
        )
        marker_colors  = [
            '#43a047' if t in ['Receipt', 'Initial'] else '#fb8c00'
            for t in df['Entry Type']
        ]
        marker_symbols = [
            'triangle-up' if t in ['Receipt', 'Initial'] else 'triangle-down'
            for t in df['Entry Type']
        ]
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=df['Date'], y=df['Closing Stock (Kg)'],
            mode='lines+markers', name='Stock',
            line=dict(color='#0f3460', width=2.5),
            fill='tozeroy', fillcolor='rgba(15,52,96,.07)',
            marker=dict(size=9, color=marker_colors, symbol=marker_symbols)
        ))
        fig.add_hline(
            y=st.session_state.low_thr,
            line_dash="dash", line_color="#e53935",
            annotation_text="Threshold: "
            + f"{st.session_state.low_thr:.0f}" + " Kg"
        )
        fig.update_layout(
            height=360, xaxis_title="Date", yaxis_title="Kg",
            hovermode='x unified',
            plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(fig, use_container_width=True)

    with ch2:
        st.markdown(
            "<div class='section-header'>📊 Receipts vs Consumption</div>",
            unsafe_allow_html=True
        )
        fig2 = go.Figure()
        if len(rcv_df) > 0:
            fig2.add_trace(go.Bar(
                x=rcv_df['Date'],
                y=rcv_df['Received from Store (MT)'],
                name='Received (MT)',
                marker_color='#43a047',
                yaxis='y2'
            ))
        if len(con_df) > 0:
            fig2.add_trace(go.Bar(
                x=con_df['Date'],
                y=con_df['Used for Sidewall Repair (Kg)'],
                name='Consumed (Kg)',
                marker_color='#fb8c00'
            ))
        fig2.update_layout(
            barmode='group', height=360,
            yaxis=dict(title='Consumed (Kg)'),
            yaxis2=dict(title='Received (MT)', overlaying='y', side='right'),
            hovermode='x unified',
            plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(fig2, use_container_width=True)

    ch3, ch4 = st.columns(2)
    with ch3:
        st.markdown(
            "<div class='section-header'>🔄 Cumulative Trends</div>",
            unsafe_allow_html=True
        )
        fig3 = go.Figure()
        if len(rcv_df) > 0:
            rs = rcv_df.sort_values('Date')
            fig3.add_trace(go.Scatter(
                x=rs['Date'],
                y=rs['Received from Store (Kg)'].cumsum(),
                mode='lines+markers', name='Cumul. Received',
                line=dict(color='#43a047', width=2)
            ))
        if len(con_df) > 0:
            ds = df.sort_values('Date')
            fig3.add_trace(go.Scatter(
                x=ds['Date'],
                y=ds['Total Consumption Till Date (Kg)'],
                mode='lines+markers', name='Cumul. Consumed',
                line=dict(color='#e53935', width=2),
                fill='tozeroy', fillcolor='rgba(229,57,53,.07)'
            ))
        fig3.update_layout(
            height=360, xaxis_title="Date", yaxis_title="Kg",
            hovermode='x unified',
            plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(fig3, use_container_width=True)

    with ch4:
        st.markdown(
            "<div class='section-header'>🥧 Stock Distribution</div>",
            unsafe_allow_html=True
        )
        fig4 = go.Figure(data=[go.Pie(
            labels=['Consumed', 'In Stock'],
            values=[tot_con, cur],
            hole=0.48,
            marker_colors=['#fb8c00', '#0f3460'],
            textinfo='label+percent+value',
            hovertemplate='%{label}: %{value:.2f} Kg (%{percent})<extra></extra>'
        )])
        fig4.update_layout(height=360, paper_bgcolor='white')
        st.plotly_chart(fig4, use_container_width=True)

    if len(con_df) > 0:
        st.markdown("---")
        st.markdown(
            "<div class='section-header'>📍 Consumption by Repair Zone</div>",
            unsafe_allow_html=True
        )
        zone_data = {}
        for z in REPAIR_ZONES:
            mask = con_df['Remarks'].str.contains(z, case=False, na=False)
            kg   = con_df.loc[mask, 'Used for Sidewall Repair (Kg)'].sum()
            if kg > 0:
                zone_data[z] = kg

        if zone_data:
            zdf  = pd.DataFrame(
                list(zone_data.items()), columns=['Zone', 'Kg Used']
            )
            zc1, zc2 = st.columns(2)
            with zc1:
                fig5 = px.bar(
                    zdf, x='Zone', y='Kg Used',
                    color='Kg Used',
                    color_continuous_scale='Oranges',
                    title="Kg Used per Zone"
                )
                fig5.update_layout(
                    height=320,
                    plot_bgcolor='white', paper_bgcolor='white'
                )
                st.plotly_chart(fig5, use_container_width=True)
            with zc2:
                fig6 = go.Figure(data=[go.Pie(
                    labels=zdf['Zone'], values=zdf['Kg Used'],
                    hole=0.4, textinfo='label+percent'
                )])
                fig6.update_layout(
                    title="Zone Share", height=320, paper_bgcolor='white'
                )
                st.plotly_chart(fig6, use_container_width=True)
        else:
            st.info("Zone data not available.")

        st.markdown(
            "<div class='section-header'>🔥 Consumption per Entry</div>",
            unsafe_allow_html=True
        )
        fig7 = px.bar(
            con_df.sort_values('Date'),
            x='Date', y='Used for Sidewall Repair (Kg)',
            color='Used for Sidewall Repair (Kg)',
            color_continuous_scale='Oranges'
        )
        fig7.add_hline(
            y=avg_use, line_dash="dot", line_color="#1e88e5",
            annotation_text="Avg: " + f"{avg_use:.0f}" + " Kg"
        )
        fig7.update_layout(
            height=320, plot_bgcolor='white', paper_bgcolor='white'
        )
        st.plotly_chart(fig7, use_container_width=True)

        st.markdown("---")
        st.markdown(
            "<div class='section-header'>📋 Statistical Summary</div>",
            unsafe_allow_html=True
        )
        s   = con_df['Used for Sidewall Repair (Kg)'].describe()
        sdf = pd.DataFrame({
            'Metric': [
                'Count', 'Mean', 'Std Dev', 'Min',
                '25th %ile', 'Median', '75th %ile', 'Max'
            ],
            'Value (Kg)': [
                f"{s['count']:.0f}", f"{s['mean']:.2f}",
                f"{s['std']:.2f}",   f"{s['min']:.2f}",
                f"{s['25%']:.2f}",   f"{s['50%']:.2f}",
                f"{s['75%']:.2f}",   f"{s['max']:.2f}"
            ]
        })
        st.dataframe(sdf, width='stretch', hide_index=True)


# ============================================================
# MODULE 8B — REPORTS (4 full reports with charts + tables + PDF/Excel)
# ============================================================

def _report_filters(key_prefix: str, df: pd.DataFrame):
    """
    Render a unified filter bar and return the filtered DataFrame.
    Filters: date range, entry type, repair zone, heat number substring.
    """
    st.markdown(
        "<div class='section-header'>🔍 Report Filters</div>",
        unsafe_allow_html=True
    )
    fc1, fc2, fc3, fc4, fc5 = st.columns(5)

    dmin  = pd.to_datetime(df['Date']).min().date()
    dmax  = pd.to_datetime(df['Date']).max().date()

    with fc1:
        dfrom = st.date_input("From:", value=dmin, key=key_prefix + "_from")
    with fc2:
        dto   = st.date_input("To:",   value=dmax, key=key_prefix + "_to")
    with fc3:
        etype_opts = ["All"] + sorted(df['Entry Type'].unique().tolist())
        etype = st.selectbox("Entry Type:", etype_opts, key=key_prefix + "_etype")
    with fc4:
        zone_opts = ["All"] + REPAIR_ZONES
        zone  = st.selectbox("Repair Zone:", zone_opts, key=key_prefix + "_zone")
    with fc5:
        heat  = st.text_input("Heat No. contains:", value="", key=key_prefix + "_heat",
                               placeholder="e.g. H-2024")

    ddf = df.copy()
    ddf['Date'] = pd.to_datetime(ddf['Date'])
    ddf = ddf[(ddf['Date'].dt.date >= dfrom) & (ddf['Date'].dt.date <= dto)]
    if etype != "All":
        ddf = ddf[ddf['Entry Type'] == etype]
    if zone != "All":
        ddf = ddf[ddf['Remarks'].str.contains(zone, case=False, na=False)]
    if heat.strip():
        ddf = ddf[ddf['Remarks'].str.contains(heat.strip(), case=False, na=False)]

    st.caption(
        f"📅 {fmt_date(pd.Timestamp(dfrom))} → {fmt_date(pd.Timestamp(dto))}  |  "
        f"**{len(ddf)}** entries after filters"
    )
    return ddf.reset_index(drop=True)


def _fig_to_image(fig) -> bytes:
    """Convert a plotly figure to PNG bytes for PDF embedding."""
    try:
        return fig.to_image(format="png", width=900, height=400, scale=2)
    except Exception:
        return None


def _build_report_excel(sheets: dict) -> bytes:
    """
    sheets = {'Sheet Name': dataframe, ...}
    Returns an Excel workbook bytes object.
    """
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for name, df_s in sheets.items():
            df_s.to_excel(writer, index=False, sheet_name=name[:31])
    return buf.getvalue()


def _build_report_pdf(title: str, sections: list) -> bytes:
    """
    sections = list of dicts:
      {'type': 'heading', 'text': str}
      {'type': 'text',    'text': str}
      {'type': 'table',   'df': DataFrame}
      {'type': 'fig',     'fig': plotly Figure}
    Returns PDF bytes using reportlab.
    """
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, Image as RLImage,
                                    PageBreak, HRFlowable)
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    import tempfile, os

    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('ReportTitle', parent=styles['Title'],
                                  fontSize=18, spaceAfter=6,
                                  textColor=colors.HexColor('#0f3460'))
    style_h2    = ParagraphStyle('H2', parent=styles['Heading2'],
                                  fontSize=13, spaceBefore=14, spaceAfter=4,
                                  textColor=colors.HexColor('#16213e'))
    style_body  = ParagraphStyle('Body', parent=styles['Normal'],
                                  fontSize=9, spaceAfter=4)
    style_meta  = ParagraphStyle('Meta', parent=styles['Normal'],
                                  fontSize=8, textColor=colors.grey)

    story = []
    story.append(Paragraph("🏭 Gunning Mass Stock Register", style_title))
    story.append(Paragraph(title, style_h2))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%d-%b-%Y %H:%M')}  |  EAF Sidewall Repair",
        style_meta
    ))
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor('#0f3460'), spaceAfter=10))

    for sec in sections:
        stype = sec.get('type')

        if stype == 'heading':
            story.append(Spacer(1, 6))
            story.append(Paragraph(sec['text'], style_h2))

        elif stype == 'text':
            story.append(Paragraph(sec['text'], style_body))

        elif stype == 'table':
            df_t = sec['df'].copy()
            df_t = df_t.fillna('')
            data = [list(df_t.columns)] + df_t.astype(str).values.tolist()

            col_w = (landscape(A4)[0] - 3*cm) / max(len(df_t.columns), 1)
            col_widths = [col_w] * len(df_t.columns)

            tbl = Table(data, colWidths=col_widths, repeatRows=1)
            tbl.setStyle(TableStyle([
                ('BACKGROUND',  (0, 0), (-1, 0),  colors.HexColor('#0f3460')),
                ('TEXTCOLOR',   (0, 0), (-1, 0),  colors.white),
                ('FONTNAME',    (0, 0), (-1, 0),  'Helvetica-Bold'),
                ('FONTSIZE',    (0, 0), (-1, 0),  8),
                ('FONTSIZE',    (0, 1), (-1, -1), 7),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1),
                 [colors.white, colors.HexColor('#f0f4ff')]),
                ('GRID',        (0, 0), (-1, -1), 0.3, colors.HexColor('#cccccc')),
                ('VALIGN',      (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING',(0, 0), (-1, -1), 4),
                ('TOPPADDING',  (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING',(0,0), (-1, -1), 3),
            ]))
            story.append(tbl)
            story.append(Spacer(1, 8))

        elif stype == 'fig':
            img_bytes = _fig_to_image(sec['fig'])
            if img_bytes:
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                tmp.write(img_bytes)
                tmp.close()
                img_w = landscape(A4)[0] - 3*cm
                img_h = img_w * 400 / 900
                story.append(RLImage(tmp.name, width=img_w, height=img_h))
                story.append(Spacer(1, 8))
                os.unlink(tmp.name)

        elif stype == 'pagebreak':
            story.append(PageBreak())

    doc.build(story)
    return buf.getvalue()


# ── Report 1: Stock Movement ──────────────────────────────────────────────────
def _render_stock_movement(ddf: pd.DataFrame, ts: str):
    st.markdown("### 📈 Stock Movement Report")

    if len(ddf) == 0:
        st.warning("No data for selected filters.")
        return

    df_plot = ddf.copy()
    df_plot['Date'] = pd.to_datetime(df_plot['Date'])
    df_plot = df_plot.sort_values('Date')

    # ── KPIs ──
    k1, k2, k3, k4 = st.columns(4)
    open_stk  = float(df_plot.iloc[0]['Opening Stock (Kg)'])
    close_stk = float(df_plot.iloc[-1]['Closing Stock (Kg)'])
    net_chg   = close_stk - open_stk
    rcv_tot   = df_plot[df_plot['Entry Type'].isin(['Receipt','Initial'])]['Received from Store (Kg)'].sum()
    con_tot   = df_plot[df_plot['Entry Type'] == 'Consumption']['Used for Sidewall Repair (Kg)'].sum()

    k1.markdown(kpi("Opening Stock",  f"{open_stk:.0f} Kg",  f"{kg2mt(open_stk):.2f} MT"),  unsafe_allow_html=True)
    k2.markdown(kpi("Closing Stock",  f"{close_stk:.0f} Kg", f"{kg2mt(close_stk):.2f} MT"), unsafe_allow_html=True)
    k3.markdown(kpi("Net Change",
                    ("+" if net_chg >= 0 else "") + f"{net_chg:.0f} Kg",
                    "received − consumed"),                                                   unsafe_allow_html=True)
    k4.markdown(kpi("Entries",        str(len(ddf)),          "in period"),                  unsafe_allow_html=True)

    st.markdown("---")

    # ── Chart 1: Stock level line ──
    marker_colors  = ['#43a047' if t in ['Receipt','Initial'] else '#fb8c00' for t in df_plot['Entry Type']]
    marker_symbols = ['triangle-up' if t in ['Receipt','Initial'] else 'triangle-down' for t in df_plot['Entry Type']]
    fig1 = go.Figure()
    fig1.add_trace(go.Scatter(
        x=df_plot['Date'], y=df_plot['Closing Stock (Kg)'],
        mode='lines+markers', name='Closing Stock',
        line=dict(color='#0f3460', width=2.5),
        fill='tozeroy', fillcolor='rgba(15,52,96,.08)',
        marker=dict(size=8, color=marker_colors, symbol=marker_symbols),
        hovertemplate='%{x|%d-%b-%Y}<br>Stock: %{y:.0f} Kg<extra></extra>'
    ))
    fig1.add_hline(y=st.session_state.low_thr, line_dash="dash",
                   line_color="#e53935",
                   annotation_text=f"Low threshold: {st.session_state.low_thr:.0f} Kg")
    fig1.update_layout(height=350, title="Stock Level Over Period",
                       xaxis_title="Date", yaxis_title="Stock (Kg)",
                       hovermode='x unified', plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 2: Daily opening vs closing bar ──
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=df_plot['Date'], y=df_plot['Opening Stock (Kg)'],
                          name='Opening', marker_color='#90caf9', opacity=0.7))
    fig2.add_trace(go.Bar(x=df_plot['Date'], y=df_plot['Closing Stock (Kg)'],
                          name='Closing', marker_color='#0f3460'))
    fig2.update_layout(barmode='group', height=320, title="Opening vs Closing Stock",
                       xaxis_title="Date", yaxis_title="Kg",
                       plot_bgcolor='white', paper_bgcolor='white')

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig1, use_container_width=True)
    with c2: st.plotly_chart(fig2, use_container_width=True)

    # ── Table ──
    st.markdown("<div class='section-header'>📋 Stock Movement Table</div>", unsafe_allow_html=True)
    tdf = df_plot[['Date','Entry Type','Opening Stock (Kg)',
                   'Received from Store (MT)','Used for Sidewall Repair (Kg)',
                   'Closing Stock (Kg)','Remarks']].copy()
    tdf['Date'] = tdf['Date'].apply(fmt_date)
    tdf = tdf.round(2)
    st.dataframe(tdf, width='stretch', hide_index=True)

    # ── Exports ──
    st.markdown("---")
    ex1, ex2 = st.columns(2)
    with ex1:
        xlsx = _build_report_excel({'Stock Movement': tdf})
        st.download_button("📥 Download Excel", data=xlsx,
                           file_name=f"Stock_Movement_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width='stretch')
    with ex2:
        try:
            pdf = _build_report_pdf("Stock Movement Report", [
                {'type': 'text',    'text': f"Period: {tdf['Date'].iloc[0]} → {tdf['Date'].iloc[-1]}  |  Entries: {len(tdf)}"},
                {'type': 'heading', 'text': 'Stock Level Chart'},
                {'type': 'fig',     'fig': fig1},
                {'type': 'heading', 'text': 'Opening vs Closing Stock'},
                {'type': 'fig',     'fig': fig2},
                {'type': 'heading', 'text': 'Detail Table'},
                {'type': 'table',   'df': tdf},
            ])
            st.download_button("📄 Download PDF", data=pdf,
                               file_name=f"Stock_Movement_{ts}.pdf",
                               mime="application/pdf", width='stretch')
        except Exception as e:
            st.warning(f"PDF export requires `kaleido` + `reportlab`. Error: {e}")


# ── Report 2: Consumption Analysis ───────────────────────────────────────────
def _render_consumption_analysis(ddf: pd.DataFrame, ts: str):
    st.markdown("### 🔥 Consumption Analysis Report")

    con_df = ddf[ddf['Entry Type'] == 'Consumption'].copy()
    if len(con_df) == 0:
        st.warning("No consumption entries for selected filters.")
        return

    con_df['Date'] = pd.to_datetime(con_df['Date'])
    con_df = con_df.sort_values('Date')

    total  = con_df['Used for Sidewall Repair (Kg)'].sum()
    avg    = con_df['Used for Sidewall Repair (Kg)'].mean()
    mx     = con_df['Used for Sidewall Repair (Kg)'].max()
    mn     = con_df['Used for Sidewall Repair (Kg)'].min()

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Total Consumed", f"{total:.0f} Kg", f"{kg2mt(total):.2f} MT"), unsafe_allow_html=True)
    k2.markdown(kpi("Avg per Entry",  f"{avg:.0f} Kg",   "per event"),              unsafe_allow_html=True)
    k3.markdown(kpi("Max Single",     f"{mx:.0f} Kg",    "highest event"),          unsafe_allow_html=True)
    k4.markdown(kpi("Events",         str(len(con_df)),  "in period"),              unsafe_allow_html=True)

    st.markdown("---")

    # ── Chart 1: Consumption per event ──
    fig1 = px.bar(con_df, x='Date', y='Used for Sidewall Repair (Kg)',
                  color='Used for Sidewall Repair (Kg)',
                  color_continuous_scale='Oranges',
                  title="Consumption per Event",
                  hover_data={'Date': '|%d-%b-%Y'})
    fig1.add_hline(y=avg, line_dash="dot", line_color="#1e88e5",
                   annotation_text=f"Avg: {avg:.0f} Kg")
    fig1.update_layout(height=320, plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 2: Zone breakdown ──
    zone_data = {}
    for z in REPAIR_ZONES:
        kg = con_df.loc[con_df['Remarks'].str.contains(z, case=False, na=False),
                        'Used for Sidewall Repair (Kg)'].sum()
        if kg > 0:
            zone_data[z] = kg

    fig2 = None
    fig3 = None
    if zone_data:
        zdf = pd.DataFrame(list(zone_data.items()), columns=['Zone', 'Kg Used'])
        fig2 = px.bar(zdf, x='Zone', y='Kg Used', color='Kg Used',
                      color_continuous_scale='Reds', title="Consumption by Repair Zone")
        fig2.update_layout(height=320, plot_bgcolor='white', paper_bgcolor='white')

        fig3 = go.Figure(data=[go.Pie(
            labels=zdf['Zone'], values=zdf['Kg Used'],
            hole=0.42, textinfo='label+percent+value',
            marker_colors=px.colors.sequential.Oranges[2:]
        )])
        fig3.update_layout(title="Zone Share", height=320, paper_bgcolor='white')

    # ── Chart 4: Monthly consumption ──
    con_df['Month'] = con_df['Date'].dt.to_period('M').astype(str)
    monthly = (con_df.groupby('Month')['Used for Sidewall Repair (Kg)']
               .agg(['sum','count','mean']).reset_index())
    monthly.columns = ['Month','Total (Kg)','Events','Avg (Kg)']
    monthly = monthly.round(2)

    fig4 = go.Figure()
    fig4.add_trace(go.Bar(x=monthly['Month'], y=monthly['Total (Kg)'],
                          name='Total (Kg)', marker_color='#fb8c00'))
    fig4.add_trace(go.Scatter(x=monthly['Month'], y=monthly['Avg (Kg)'],
                              mode='lines+markers', name='Avg (Kg)',
                              line=dict(color='#1e88e5', width=2), yaxis='y2'))
    fig4.update_layout(
        barmode='group', height=320, title="Monthly Consumption",
        yaxis=dict(title='Total (Kg)'),
        yaxis2=dict(title='Avg (Kg)', overlaying='y', side='right'),
        plot_bgcolor='white', paper_bgcolor='white'
    )

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig1, use_container_width=True)
    with c2: st.plotly_chart(fig4, use_container_width=True)

    if fig2 and fig3:
        c3, c4 = st.columns(2)
        with c3: st.plotly_chart(fig2, use_container_width=True)
        with c4: st.plotly_chart(fig3, use_container_width=True)

    # ── Tables ──
    st.markdown("<div class='section-header'>📋 Consumption Detail Table</div>", unsafe_allow_html=True)
    tdf = con_df[['Date','Opening Stock (Kg)','Used for Sidewall Repair (Kg)',
                  'Closing Stock (Kg)','Total Consumption Till Date (Kg)','Remarks']].copy()
    tdf['Date'] = tdf['Date'].apply(fmt_date)
    tdf = tdf.round(2)
    st.dataframe(tdf, width='stretch', hide_index=True)

    if zone_data:
        st.markdown("<div class='section-header'>📍 Zone Summary</div>", unsafe_allow_html=True)
        st.dataframe(zdf.sort_values('Kg Used', ascending=False).round(2),
                     width='stretch', hide_index=True)

    st.markdown("<div class='section-header'>📅 Monthly Summary</div>", unsafe_allow_html=True)
    st.dataframe(monthly, width='stretch', hide_index=True)

    # ── Exports ──
    st.markdown("---")
    ex1, ex2 = st.columns(2)
    with ex1:
        sheets = {'Consumption Detail': tdf, 'Monthly Summary': monthly}
        if zone_data:
            sheets['Zone Breakdown'] = zdf.round(2)
        xlsx = _build_report_excel(sheets)
        st.download_button("📥 Download Excel", data=xlsx,
                           file_name=f"Consumption_Analysis_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width='stretch')
    with ex2:
        try:
            sections = [
                {'type': 'text',    'text': f"Total: {total:.0f} Kg  |  Avg: {avg:.0f} Kg  |  Events: {len(con_df)}"},
                {'type': 'heading', 'text': 'Consumption per Event'},
                {'type': 'fig',     'fig': fig1},
                {'type': 'heading', 'text': 'Monthly Consumption'},
                {'type': 'fig',     'fig': fig4},
            ]
            if fig2:
                sections += [
                    {'type': 'heading', 'text': 'Consumption by Zone'},
                    {'type': 'fig',     'fig': fig2},
                    {'type': 'fig',     'fig': fig3},
                ]
            sections += [
                {'type': 'heading', 'text': 'Detail Table'},
                {'type': 'table',   'df': tdf},
                {'type': 'heading', 'text': 'Monthly Summary'},
                {'type': 'table',   'df': monthly},
            ]
            pdf = _build_report_pdf("Consumption Analysis Report", sections)
            st.download_button("📄 Download PDF", data=pdf,
                               file_name=f"Consumption_Analysis_{ts}.pdf",
                               mime="application/pdf", width='stretch')
        except Exception as e:
            st.warning(f"PDF export requires `kaleido` + `reportlab`. Error: {e}")


# ── Report 3: Receipt / Procurement ──────────────────────────────────────────
def _render_receipt_report(ddf: pd.DataFrame, ts: str):
    st.markdown("### 📥 Receipt / Procurement Report")

    rcv_df = ddf[ddf['Entry Type'].isin(['Receipt','Initial'])].copy()
    if len(rcv_df) == 0:
        st.warning("No receipt entries for selected filters.")
        return

    rcv_df['Date'] = pd.to_datetime(rcv_df['Date'])
    rcv_df = rcv_df.sort_values('Date')

    total_kg = rcv_df['Received from Store (Kg)'].sum()
    total_mt = rcv_df['Received from Store (MT)'].sum()
    avg_mt   = rcv_df['Received from Store (MT)'].mean()

    k1, k2, k3 = st.columns(3)
    k1.markdown(kpi("Total Received", f"{total_kg:.0f} Kg", f"{total_mt:.1f} MT"), unsafe_allow_html=True)
    k2.markdown(kpi("Avg per Receipt", f"{avg_mt:.1f} MT",   "per GRN"),            unsafe_allow_html=True)
    k3.markdown(kpi("Receipts",        str(len(rcv_df)),      "in period"),          unsafe_allow_html=True)

    st.markdown("---")

    # ── Chart 1: Receipts over time ──
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=rcv_df['Date'], y=rcv_df['Received from Store (MT)'],
                          name='Received (MT)', marker_color='#43a047',
                          hovertemplate='%{x|%d-%b-%Y}<br>%{y:.1f} MT<extra></extra>'))
    fig1.update_layout(height=320, title="Receipts Over Time",
                       xaxis_title="Date", yaxis_title="MT",
                       plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 2: Cumulative received ──
    rcv_df['Cumul Received (Kg)'] = rcv_df['Received from Store (Kg)'].cumsum()
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(x=rcv_df['Date'], y=rcv_df['Cumul Received (Kg)'],
                              mode='lines+markers', name='Cumulative',
                              line=dict(color='#43a047', width=2.5),
                              fill='tozeroy', fillcolor='rgba(67,160,71,.1)'))
    fig2.update_layout(height=320, title="Cumulative Receipts",
                       xaxis_title="Date", yaxis_title="Kg",
                       plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 3: Monthly receipts ──
    rcv_df['Month'] = rcv_df['Date'].dt.to_period('M').astype(str)
    monthly = (rcv_df.groupby('Month')['Received from Store (Kg)']
               .agg(['sum','count']).reset_index())
    monthly.columns = ['Month','Total Received (Kg)','No. of Receipts']
    monthly['Total Received (MT)'] = (monthly['Total Received (Kg)'] / 1000).round(3)
    monthly = monthly.round(2)

    fig3 = px.bar(monthly, x='Month', y='Total Received (Kg)',
                  color='Total Received (Kg)', color_continuous_scale='Greens',
                  title="Monthly Receipts (Kg)")
    fig3.update_layout(height=300, plot_bgcolor='white', paper_bgcolor='white')

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig1, use_container_width=True)
    with c2: st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)

    # ── Tables ──
    st.markdown("<div class='section-header'>📋 Receipt Detail Table</div>", unsafe_allow_html=True)
    tdf = rcv_df[['Date','Entry Type','Received from Store (MT)','Received from Store (Kg)',
                  'Closing Stock (Kg)','Remarks']].copy()
    tdf['Date'] = tdf['Date'].apply(fmt_date)
    tdf = tdf.round(2)
    st.dataframe(tdf, width='stretch', hide_index=True)

    st.markdown("<div class='section-header'>📅 Monthly Summary</div>", unsafe_allow_html=True)
    st.dataframe(monthly, width='stretch', hide_index=True)

    # ── Exports ──
    st.markdown("---")
    ex1, ex2 = st.columns(2)
    with ex1:
        xlsx = _build_report_excel({'Receipt Detail': tdf, 'Monthly Summary': monthly})
        st.download_button("📥 Download Excel", data=xlsx,
                           file_name=f"Receipt_Report_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width='stretch')
    with ex2:
        try:
            pdf = _build_report_pdf("Receipt / Procurement Report", [
                {'type': 'text',    'text': f"Total: {total_kg:.0f} Kg ({total_mt:.1f} MT)  |  Receipts: {len(rcv_df)}"},
                {'type': 'heading', 'text': 'Receipts Over Time'},
                {'type': 'fig',     'fig': fig1},
                {'type': 'heading', 'text': 'Cumulative Receipts'},
                {'type': 'fig',     'fig': fig2},
                {'type': 'heading', 'text': 'Monthly Receipts'},
                {'type': 'fig',     'fig': fig3},
                {'type': 'heading', 'text': 'Detail Table'},
                {'type': 'table',   'df': tdf},
                {'type': 'heading', 'text': 'Monthly Summary'},
                {'type': 'table',   'df': monthly},
            ])
            st.download_button("📄 Download PDF", data=pdf,
                               file_name=f"Receipt_Report_{ts}.pdf",
                               mime="application/pdf", width='stretch')
        except Exception as e:
            st.warning(f"PDF export requires `kaleido` + `reportlab`. Error: {e}")


# ── Report 4: Monthly Summary ─────────────────────────────────────────────────
def _render_monthly_summary(ddf: pd.DataFrame, ts: str):
    st.markdown("### 📅 Monthly Summary Report")

    if len(ddf) == 0:
        st.warning("No data for selected filters.")
        return

    df_m = ddf.copy()
    df_m['Date']  = pd.to_datetime(df_m['Date'])
    df_m['Month'] = df_m['Date'].dt.to_period('M').astype(str)

    # Build monthly agg
    con_m = (df_m[df_m['Entry Type'] == 'Consumption']
             .groupby('Month')['Used for Sidewall Repair (Kg)']
             .agg(Con_Total='sum', Con_Events='count', Con_Avg='mean', Con_Max='max')
             .reset_index())
    rcv_m = (df_m[df_m['Entry Type'].isin(['Receipt','Initial'])]
             .groupby('Month')['Received from Store (Kg)']
             .agg(Rcv_Total='sum', Rcv_Events='count')
             .reset_index())
    merged = pd.merge(rcv_m, con_m, on='Month', how='outer').fillna(0).sort_values('Month')
    merged['Rcv_MT']    = (merged['Rcv_Total'] / 1000).round(3)
    merged['Net Change (Kg)'] = (merged['Rcv_Total'] - merged['Con_Total']).round(2)
    merged = merged.round(2)
    merged.columns = ['Month','Received (Kg)','Receipts','Consumed (Kg)',
                      'Con. Events','Avg Con. (Kg)','Max Con. (Kg)',
                      'Received (MT)','Net Change (Kg)']

    # KPIs
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(kpi("Months",         str(len(merged)),                               "in period"),         unsafe_allow_html=True)
    k2.markdown(kpi("Total Received", f"{merged['Received (Kg)'].sum():.0f} Kg",      f"{merged['Received (MT)'].sum():.1f} MT"), unsafe_allow_html=True)
    k3.markdown(kpi("Total Consumed", f"{merged['Consumed (Kg)'].sum():.0f} Kg",      ""),                  unsafe_allow_html=True)
    k4.markdown(kpi("Net Change",     f"{merged['Net Change (Kg)'].sum():+.0f} Kg",   "rcv − con"),         unsafe_allow_html=True)

    st.markdown("---")

    # ── Chart 1: Receipt vs Consumption by month ──
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=merged['Month'], y=merged['Received (Kg)'],
                          name='Received (Kg)', marker_color='#43a047', opacity=0.85))
    fig1.add_trace(go.Bar(x=merged['Month'], y=merged['Consumed (Kg)'],
                          name='Consumed (Kg)', marker_color='#fb8c00', opacity=0.85))
    fig1.update_layout(barmode='group', height=340,
                       title="Monthly Received vs Consumed",
                       xaxis_title="Month", yaxis_title="Kg",
                       plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 2: Net change waterfall ──
    fig2 = go.Figure(go.Waterfall(
        name="Net Change",
        orientation="v",
        x=merged['Month'],
        y=merged['Net Change (Kg)'],
        connector={"line": {"color": "#aaa"}},
        increasing={"marker": {"color": "#43a047"}},
        decreasing={"marker": {"color": "#e53935"}},
    ))
    fig2.update_layout(height=320, title="Monthly Net Stock Change (Received − Consumed)",
                       xaxis_title="Month", yaxis_title="Kg",
                       plot_bgcolor='white', paper_bgcolor='white')

    # ── Chart 3: Consumption trend line ──
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(x=merged['Month'], y=merged['Consumed (Kg)'],
                              mode='lines+markers+text',
                              text=merged['Consumed (Kg)'].astype(int).astype(str),
                              textposition='top center',
                              line=dict(color='#fb8c00', width=2.5),
                              name='Consumed (Kg)'))
    fig3.add_trace(go.Scatter(x=merged['Month'], y=merged['Avg Con. (Kg)'],
                              mode='lines', name='Avg per Event',
                              line=dict(color='#1e88e5', width=1.5, dash='dot')))
    fig3.update_layout(height=300, title="Monthly Consumption Trend",
                       plot_bgcolor='white', paper_bgcolor='white')

    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(fig1, use_container_width=True)
    with c2: st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)

    # ── Table ──
    st.markdown("<div class='section-header'>📋 Monthly Summary Table</div>", unsafe_allow_html=True)
    st.dataframe(merged, width='stretch', hide_index=True)

    # ── Exports ──
    st.markdown("---")
    ex1, ex2 = st.columns(2)
    with ex1:
        xlsx = _build_report_excel({'Monthly Summary': merged})
        st.download_button("📥 Download Excel", data=xlsx,
                           file_name=f"Monthly_Summary_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width='stretch')
    with ex2:
        try:
            pdf = _build_report_pdf("Monthly Summary Report", [
                {'type': 'heading', 'text': 'Received vs Consumed by Month'},
                {'type': 'fig',     'fig': fig1},
                {'type': 'heading', 'text': 'Net Stock Change (Waterfall)'},
                {'type': 'fig',     'fig': fig2},
                {'type': 'heading', 'text': 'Consumption Trend'},
                {'type': 'fig',     'fig': fig3},
                {'type': 'heading', 'text': 'Monthly Summary Table'},
                {'type': 'table',   'df': merged},
            ])
            st.download_button("📄 Download PDF", data=pdf,
                               file_name=f"Monthly_Summary_{ts}.pdf",
                               mime="application/pdf", width='stretch')
        except Exception as e:
            st.warning(f"PDF export requires `kaleido` + `reportlab`. Error: {e}")


# ── Main Reports renderer ─────────────────────────────────────────────────────
def render_reports():
    st.header("📋 Reports")
    if len(st.session_state.stock_data) == 0:
        st.warning("No data available.")
        return

    df = st.session_state.stock_data.copy()
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Shared filter bar
    ddf = _report_filters("rpt", df)

    st.markdown("---")

    # Tab per report
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Stock Movement",
        "🔥 Consumption Analysis",
        "📥 Receipt / Procurement",
        "📅 Monthly Summary",
    ])
    with tab1: _render_stock_movement(ddf, ts)
    with tab2: _render_consumption_analysis(ddf, ts)
    with tab3: _render_receipt_report(ddf, ts)
    with tab4: _render_monthly_summary(ddf, ts)


# ============================================================
# MODULE 8 — DOWNLOAD DATA (renamed, kept for raw file exports)
# ============================================================
def render_download_reports():
    st.header("💾 Download Raw Data")
    st.markdown(
        "<div class='section-header'>"
        "📁 Download raw CSV / Excel files. For charts and analysis go to 📋 Reports."
        "</div>",
        unsafe_allow_html=True
    )

    if len(st.session_state.stock_data) == 0:
        st.warning("No data available to export.")
        return

    df       = st.session_state.stock_data.copy()
    df['Date'] = pd.to_datetime(df['Date'])
    ts       = datetime.now().strftime('%Y%m%d_%H%M%S')

    st.markdown(
        "<div class='section-header'>🔍 Date Range Filter</div>",
        unsafe_allow_html=True
    )
    dr1, dr2 = st.columns(2)
    with dr1:
        dmin  = df['Date'].min().date()
        dfrom = st.date_input("From:", value=dmin, key="rep_from")
    with dr2:
        dmax  = df['Date'].max().date()
        dto   = st.date_input("To:",   value=dmax, key="rep_to")

    df_filt  = df[(df['Date'].dt.date >= dfrom) & (df['Date'].dt.date <= dto)]
    con_filt = df_filt[df_filt['Entry Type'] == 'Consumption']
    rcv_filt = df_filt[df_filt['Entry Type'].isin(['Receipt', 'Initial'])]

    st.info(
        "📅 Filtered range: **" + fmt_date(pd.Timestamp(dfrom))
        + "** → **" + fmt_date(pd.Timestamp(dto)) + "** |  "
        + "**" + str(len(df_filt)) + "** total entries  |  "
        + "**" + str(len(rcv_filt)) + "** receipts  |  "
        + "**" + str(len(con_filt)) + "** consumptions"
    )

    st.markdown("---")

    # Report 1
    st.markdown(
        "<div class='download-card'><h3>📋 Report 1 — Full Stock Register</h3>"
        "<p>All entries with a Summary sheet.</p></div>",
        unsafe_allow_html=True
    )
    with st.expander("👁️ Preview", expanded=False):
        st.dataframe(prepare_export_df(df_filt).head(10),
                     width='stretch', hide_index=True)
    download_widget("📥 Download Full Register (CSV)",
                    build_csv(prepare_export_df(df_filt)),
                    "Full_Stock_Register_" + ts, "text/csv", ".csv")

    st.markdown("---")

    # Report 2
    st.markdown(
        "<div class='download-card'><h3>🔥 Report 2 — Consumption Report</h3>"
        "<p>Consumption entries with Zone Breakdown and Monthly Summary.</p></div>",
        unsafe_allow_html=True
    )
    if len(con_filt) == 0:
        st.info("No consumption entries in the selected date range.")
    else:
        total_used = con_filt['Used for Sidewall Repair (Kg)'].sum()
        avg_used   = con_filt['Used for Sidewall Repair (Kg)'].mean()
        max_used   = con_filt['Used for Sidewall Repair (Kg)'].max()
        ks1, ks2, ks3, ks4 = st.columns(4)
        ks1.markdown(kpi("Entries", str(len(con_filt))), unsafe_allow_html=True)
        ks2.markdown(kpi("Total Used", f"{total_used:.0f} Kg", f"{kg2mt(total_used):.2f} MT"), unsafe_allow_html=True)
        ks3.markdown(kpi("Avg per Entry", f"{avg_used:.0f} Kg"), unsafe_allow_html=True)
        ks4.markdown(kpi("Max Single", f"{max_used:.0f} Kg"), unsafe_allow_html=True)
        download_widget("📥 Download Consumption (CSV)",
                        build_csv(prepare_export_df(con_filt)),
                        "Consumption_Report_" + ts, "text/csv", ".csv")

    st.markdown("---")

    # Report 3
    st.markdown(
        "<div class='download-card'><h3>📥 Report 3 — Receipt Report</h3>"
        "<p>Receipt entries with Monthly Summary.</p></div>",
        unsafe_allow_html=True
    )
    if len(rcv_filt) == 0:
        st.info("No receipt entries in the selected date range.")
    else:
        total_rcv_kg = rcv_filt['Received from Store (Kg)'].sum()
        total_rcv_mt = rcv_filt['Received from Store (MT)'].sum()
        rs1, rs2, rs3 = st.columns(3)
        rs1.markdown(kpi("Receipts", str(len(rcv_filt))), unsafe_allow_html=True)
        rs2.markdown(kpi("Total Received", f"{total_rcv_kg:.0f} Kg"), unsafe_allow_html=True)
        rs3.markdown(kpi("Total Received", f"{total_rcv_mt:.1f} MT"), unsafe_allow_html=True)
        download_widget("📥 Download Receipts (CSV)",
                        build_csv(prepare_export_df(rcv_filt)),
                        "Receipt_Report_" + ts, "text/csv", ".csv")

    st.markdown("---")

    # Report 4
    st.markdown(
        "<div class='download-card'><h3>📅 Report 4 — Monthly Summary</h3>"
        "<p>Month-wise aggregation of receipts and consumption.</p></div>",
        unsafe_allow_html=True
    )
    df_filt2 = df_filt.copy()
    df_filt2['Month'] = df_filt2['Date'].dt.to_period('M').astype(str)

    con_monthly = (
        df_filt2[df_filt2['Entry Type'] == 'Consumption']
        .groupby('Month')['Used for Sidewall Repair (Kg)']
        .agg(['sum', 'count', 'mean', 'max'])
        .reset_index()
    )
    con_monthly.columns = ['Month', 'Total Consumed (Kg)', 'No. of Entries',
                            'Avg per Entry (Kg)', 'Max Single (Kg)']
    con_monthly['Total Consumed (MT)'] = (con_monthly['Total Consumed (Kg)'] / 1000).round(3)
    con_monthly = con_monthly.round(2)

    rcv_monthly = (
        df_filt2[df_filt2['Entry Type'].isin(['Receipt', 'Initial'])]
        .groupby('Month')['Received from Store (Kg)']
        .agg(['sum', 'count'])
        .reset_index()
    )
    rcv_monthly.columns = ['Month', 'Total Received (Kg)', 'No. of Receipts']
    rcv_monthly['Total Received (MT)'] = (rcv_monthly['Total Received (Kg)'] / 1000).round(3)

    if len(con_monthly) > 0 or len(rcv_monthly) > 0:
        with st.expander("👁️ Preview — Monthly Summary", expanded=False):
            st.markdown("**Consumption by Month:**")
            st.dataframe(con_monthly, width='stretch', hide_index=True)
            st.markdown("**Receipts by Month:**")
            st.dataframe(rcv_monthly, width='stretch', hide_index=True)

        def build_monthly_excel() -> bytes:
            buf2 = BytesIO()
            with pd.ExcelWriter(buf2, engine='openpyxl') as writer:
                if len(con_monthly) > 0:
                    con_monthly.to_excel(writer, index=False, sheet_name='Monthly Consumption')
                if len(rcv_monthly) > 0:
                    rcv_monthly.to_excel(writer, index=False, sheet_name='Monthly Receipts')
            return buf2.getvalue()

        mc1, mc2 = st.columns(2)
        with mc1:
            download_widget("📥 Download Monthly Consumption (CSV)",
                            build_csv(con_monthly),
                            "Monthly_Consumption_Summary_" + ts, "text/csv", ".csv")
        with mc2:
            download_widget("📥 Download Monthly Receipts (CSV)",
                            build_csv(rcv_monthly),
                            "Monthly_Receipts_Summary_" + ts, "text/csv", ".csv")
    else:
        st.info("No data available for monthly summary in the selected range.")


# ============================================================
# MODULE 9 — IMPORT DATA
# ============================================================
def render_import():
    st.header("📤 Import Data")
    st.warning(
        "⚠️ Importing will **replace ALL current data**. "
        "Please download a backup first."
    )

    uploaded = st.file_uploader("Upload CSV file:", type=['csv'])
    if uploaded:
        try:
            imp_df = pd.read_csv(uploaded)
            imp_df['Date'] = pd.to_datetime(imp_df['Date'])
            missing = set(STOCK_COLUMNS) - set(imp_df.columns)
            if missing:
                st.error("❌ Missing columns: " + str(missing))
            else:
                st.success("✅ File structure validated!")
                st.dataframe(imp_df.head(5), width='stretch', hide_index=True)
                st.info("**" + str(len(imp_df)) + " records** found in file.")
                if st.button(
                    "🔄 Import & Replace Data",
                    type="primary",
                    width='stretch'
                ):
                    st.session_state.stock_data        = imp_df
                    st.session_state.initial_stock_set = True
                    save_data(imp_df)
                    st.success("✅ Data imported successfully!")
                    st.rerun()
        except Exception as e:
            st.error("❌ Error reading file: " + str(e))


# ============================================================
# MAIN ROUTER
# ============================================================
if not is_init() or len(st.session_state.stock_data) == 0:
    render_initial_setup()
else:
    if   action == "🏠 Dashboard":        render_dashboard()
    elif action == "📥 Receive Stock":     render_receive()
    elif action == "🔥 Log Consumption":   render_consumption()
    elif action == "📖 View Register":     render_register()
    elif action == "✏️ Edit Entry":        render_edit()
    elif action == "📊 Analytics":         render_analytics()
    elif action == "📋 Reports":           render_reports()
    elif action == "💾 Download Data":     render_download_reports()
    elif action == "📤 Import Data":       render_import()


# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#aaa;font-size:12px;padding:6px 0;'>"
    "🏭 Gunning Mass Stock Management System &nbsp;|&nbsp; "
    "EAF Sidewall Repair &nbsp;|&nbsp; "
    "Data saved to <code>Google Sheets (StockData tab)</code>"
    "</div>",
    unsafe_allow_html=True
)
