import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Increase Pandas Styler limit for large datasets
pd.set_option("styler.render.max_elements", 2000000)

# Set page configuration
st.set_page_config(
    page_title="Budget Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS for premium look
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .metric-container {
        display: flex;
        justify-content: space-between;
        gap: 20px;
        margin-bottom: 30px;
    }
    h1, h2, h3 {
        color: #1e293b;
        font-family: 'Inter', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# --- Multi-language Support ---
if "lang" not in st.session_state:
    st.session_state.lang = "Tiếng Việt"

st.sidebar.title("🌐 Language / Ngôn ngữ")
selected_lang = st.sidebar.radio("Select Language:", options=["Korean", "Tiếng Việt"], index=1, label_visibility="collapsed")
st.session_state.lang = selected_lang

# Translation Dictionary
LANG = {
    "Korean": {
        "filters": "🔍 Filters",
        "load_filter": "📂 Load saved filter:",
        "select_cols": "Select columns to filter by:",
        "view_month": "📅 View by month:", # Keep month selector as is or translate? Let's use Korean for consistency.
        "all_year": "Cả năm (Tổng)",
        "currency": "💵 Currency Options",
        "display_currency": "Select Display Currency",
        "ex_rate": "Exchange Rate (Divides Base Value)",
        "ex_rate_cap": "Input rate depending on your base data",
        "save_preset": "💾 Save Filter Preset",
        "preset_name": "Preset name:",
        "save_btn": "Save Preset",
        "total_budget": "💰 Total Budget",
        "total_actual": "💸 Total Actual Expense",
        "usage_pct": "📈 Total Usage %",
        "chart_overview": "Budget vs Actual Overview",
        "chart_ratio": "Overall Usage Ratio",
        "gl_detail": "📂 GL Account Detail",
        "over_budget_title": "🚨 Over Budget Warning: Detailed Accounts",
        "over_budget_cap": "Drill-down breakdown for Account Groups exceeding 100% target usage.",
        "detailed_report": "Detailed Expense Report",
        "gt": "GT",
        "dept": "Department",
        "leader": "Leader",
        "gl": "GL Account",
        "acc_id": "Detail Account ID",
        "acc_name": "Detail Account Name",
        "target": "Target Budget",
        "increase": "Increase",
        "early": "Early & Carried",
        "result": "Result",
        "usage_tar": "Usage(%) (Target)",
        "usage_all": "Usage(%) (All)"
    },
    "Tiếng Việt": {
        "filters": "🔍 Filters",
        "load_filter": "📂 Load saved filter:",
        "select_cols": "Select columns to filter:",
        "view_month": "📅 View by month:",
        "all_year": "Full Year (Total)",
        "currency": "💵 Currency Options",
        "display_currency": "Select Currency",
        "ex_rate": "Exchange Rate",
        "ex_rate_cap": "Input rate (divides base value)",
        "save_preset": "💾 Save Current Filter",
        "preset_name": "Preset name:",
        "save_btn": "Save Preset",
        "total_budget": "💰 Total Budget",
        "total_actual": "💸 Total Actual",
        "usage_pct": "📈 Total Usage %",
        "chart_overview": "Budget vs Actual Overview",
        "chart_ratio": "Usage Ratio",
        "gl_detail": "📂 Detail by Account Group",
        "over_budget_title": "🚨 Over Budget Warning: Detailed Accounts",
        "over_budget_cap": "Drill-down for groups exceeding 100% budget.",
        "detailed_report": "Detailed Expense Report",
        "gt": "GT Group",
        "dept": "Department",
        "leader": "Manager",
        "gl": "Account Group (GL)",
        "acc_id": "Account ID",
        "acc_name": "Account Name",
        "target": "Target Budget",
        "increase": "Additional Budget",
        "early": "Early & Carried",
        "result": "Actual Result",
        "usage_tar": "Usage % (Target)",
        "usage_all": "Usage % (Total)"
    }
}
L = LANG[st.session_state.lang]

# Content Translation Map (Korean -> Vietnamese)
GL_VALUE_MAP = {
    "급료와임금": "Lương & Tiền công",
    "수선비": "Chi phí sửa chữa",
    "여비교통비": "Chi phí đi lại",
    "복리후생비": "Chi phí phúc lợi",
    "보험료": "Phí bảo hiểm",
    "통신비": "Chi phí viễn thông",
    "지급임차료": "Chi phí thuê nhà",
    "세금과공과": "Thuế & Lệ phí",
    "교육훈련비": "Chi phí đào tạo",
    "차량유지비": "Chi phí xe cộ",
    "지급수수료": "Phí dịch vụ",
    "소모품비": "Chi phí vật tư",
    "사무용품비": "Chi phí văn phòng phẩm",
    "회의비": "Chi phí hội nghị",
    "운반비": "Phí vận chuyển",
    "원재료비": "Chi phí nguyên liệu",
    "비용청구서": "Hóa đơn thanh toán",
    "수도광열비": "Chi phí điện nước",
    "Meet관리": "Phí quản lý Meet",
    "감가상각비": "Khấu hao tài sản cố định",
    "도서인쇄비": "Chi phí sách & In ấn",
    "무형자산상각비": "Khấu hao tài sản vô hình",
    "접대비": "Chi phí tiếp khách",
    "수선비(공구와기구)": "Chi phí sửa chữa (Công cụ & Dụng cụ)",
    "여비교통비(해외출장-교통비)": "Chi phí đi lại (Công tác nước ngoài-Phí đi lại)",
    "고정": "Cố định",
    "변동": "Biến đổi",
    "통제 계정": "Tài khoản kiểm soát",
    "비통제 계정": "Tài khoản không kiểm soát"
}

def translate_content(df_to_trans):
    if st.session_state.lang != "Tiếng Việt":
        return df_to_trans
    
    df_new = df_to_trans.copy()
    # Find columns that might need translation (IDed by L values or original names)
    target_cols = [L["gl"], L["acc_name"], "GL Account", "Detail Account Name"]
    for col in target_cols:
        if col in df_new.columns:
            # We use a fuzzy match or simple replace for known terms
            for kor, vie in GL_VALUE_MAP.items():
                df_new[col] = df_new[col].astype(str).str.replace(kor, vie, regex=False)
    return df_new


@st.cache_data
def load_data(file_input):
    # Check if input is a path or an uploaded file object
    if isinstance(file_input, str):
        if not os.path.exists(file_input):
            return None
        df = pd.read_excel(file_input)
    else:
        # It's an uploaded file object
        df = pd.read_excel(file_input)
    
    # Auto-detect long format (pivot needed)
    c_class = None
    c_amount = None
    
    # Let's find the column that contains budget descriptors
    for col in df.columns:
        if df[col].dtype == 'object' or pd.api.types.is_string_dtype(df[col]):
            sample_vals = set(df[col].dropna().head(1000).astype(str))
            if any('1.목표예산' in v or '4.비용실적' in v for v in sample_vals):
                c_class = col
                break
                
    if c_class:
        # Identify month columns like "1월 금액", "2월 금액" ... "12월 금액"
        import re as _re
        month_pat = _re.compile(r'^(1[0-2]|[1-9])월 금액$')
        month_amount_cols = [c for c in df.columns if month_pat.match(str(c).strip())]
        
        # Total amount col = has 금액 but is NOT a month column
        total_amount_cands = [c for c in df.columns if '금액' in str(c) and c not in month_amount_cols]
        if total_amount_cands:
            c_amount = total_amount_cands[0]
        else:
            idx = df.columns.get_loc(c_class)
            c_amount = df.columns[idx + 1] if idx + 1 < len(df.columns) else None
            
        if c_amount:
            # Dimension columns ONLY - exclude class, total amount, AND month amounts
            dim_cols = [c for c in df.columns if c not in [c_class, c_amount] + month_amount_cols]
            
            # --- ROOT CAUSE FIX ---
            # Pandas pivot/merge drops rows if dimension columns have NaN.
            # Filling NaNs with empty strings ("") preserves 100% of rows perfectly.
            df[dim_cols] = df[dim_cols].fillna("")

            # STEP 1: Extract monthly data for EACH budget type before pivot
            all_monthly_data = {}
            if month_amount_cols:
                for btype in df[c_class].unique():
                    if pd.isna(btype): continue
                    btype_rows = df[df[c_class] == btype]
                    btype_monthly = btype_rows.groupby(dim_cols)[month_amount_cols].sum().reset_index()
                    rename_map = {mc: f"{mc}_{btype}" for mc in month_amount_cols}
                    btype_monthly = btype_monthly.rename(columns=rename_map)
                    all_monthly_data[btype] = btype_monthly

            # STEP 2: Pivot annual totals only
            df_wide = pd.pivot_table(
                df,
                index=dim_cols,
                columns=c_class,
                values=c_amount,
                aggfunc='sum',
                fill_value=0
            ).reset_index()
            df_wide.columns.name = None

            # STEP 3: Merge monthly data back for each budget type
            df_result = df_wide
            for btype, btype_monthly in all_monthly_data.items():
                merge_cols = [c for c in dim_cols if c in btype_monthly.columns]
                new_typed_cols = [c for c in btype_monthly.columns if c not in merge_cols]
                df_result = pd.merge(df_result, btype_monthly, on=merge_cols, how='left')
                for tc in new_typed_cols:
                    if tc in df_result.columns:
                        df_result[tc] = df_result[tc].fillna(0)
            df = df_result

    header_cols = df.columns.tolist()
    
    # Robust column detection helper
    def find_col(prefixes):
        for c in header_cols:
            c_str = str(c).strip()
            if any(c_str.startswith(p) for p in prefixes):
                return c
        return None

    c_budget = find_col(["1.목표예산", "Target Budget"])
    c_increase = find_col(["2.증액예산", "Increase"])
    c_early = find_col(["3.조기 및 이월예산", "Early"])
    c_actual = find_col(["4.비용실적", "Result", "Actual"])
    
    # Ensure numeric types for calculation columns
    calc_cols = [c for c in [c_budget, c_increase, c_early, c_actual] if c]
    for col in calc_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # --- FIX: Recompute annual totals for c_increase and c_early from monthly sums ---
    # In some data formats the annual summary column is 0 and actual amounts
    # are only stored in the per-month columns, so we rebuild the annual column.
    import re as _re2
    for annual_col in [c_increase, c_early]:
        if annual_col and annual_col in df.columns:
            # Find all monthly columns that belong to this budget type
            monthly_siblings = [
                c for c in df.columns
                if _re2.match(r'^(1[0-2]|[1-9])월 금액_', str(c))
                and str(c).endswith(f"_{annual_col}")
            ]
            if monthly_siblings:
                # Sum the monthly columns to get the correct annual total
                df[annual_col] = sum(
                    pd.to_numeric(df[mc], errors='coerce').fillna(0)
                    for mc in monthly_siblings
                )

    # Calculate Usage (%) Target = Actual / Target Budget
    if c_actual and c_budget:
        df["Usage(%) Target"] = (df[c_actual] / df[c_budget]).replace([float('inf'), -float('inf')], 0).fillna(0)
    
    # Calculate Usage (%) All = Actual / (Target + Increase + Early)
    if all([c_actual, c_budget, c_increase, c_early]):
        total_budget = df[c_budget] + df[c_increase] + df[c_early]
        df["Usage(%) All"] = (df[c_actual] / total_budget).replace([float('inf'), -float('inf')], 0).fillna(0)
    
    return df

# Helper for styling
def get_styled_df(df_to_style, currency_sym="$"):
    def color_usage(val):
        color = 'red' if val > 1.0 else 'green'
        return f'color: {color}; font-weight: bold'
    
    # Define columns to format (original, renamed, and TRANSLATED)
    # Use dict.fromkeys to remove duplicates if translated labels match English names
    money_cols = list(dict.fromkeys([
        "1.목표예산", "2.증액예산", "3.조기 및 이월예산", "4.비용실적",
        "Target Budget", "Increase", "Early & Carried", "Result",
        L["target"], L["increase"], L["early"], L["result"]
    ]))
    pct_cols = list(dict.fromkeys([
        "Usage(%) Target", "Usage(%) All", 
        "Usage(%) (Target)", "Usage(%) (All)",
        L["usage_tar"], L["usage_all"]
    ]))
    
    format_dict = {}
    for col in money_cols:
        if col in df_to_style.columns:
            if currency_sym in ["₫", "VND", "₩", "KRW"]:
                format_dict[col] = f"{currency_sym} {{:,.0f}}"
            else:
                format_dict[col] = f"{currency_sym}{{:.2f}}"
    for col in pct_cols:
        if col in df_to_style.columns:
            format_dict[col] = "{:.1%}"

    styled = df_to_style.style.format(format_dict)
    
    # Apply color to percentage columns if they exist
    existing_pct_cols = [c for c in pct_cols if c in df_to_style.columns]
    if existing_pct_cols:
        styled = styled.applymap(color_usage, subset=existing_pct_cols)
        
    return styled

# Main Logic
st.title("📊 Financial Budget Dashboard")
st.markdown("---")

# File Path
# File Path
FILE_PATH = "data.xlsx"
import datetime

# --- Upload Logic ---
st.sidebar.markdown("---")
# 1. Personal View (Temporary)
st.sidebar.header("📁 Personal Data View")
st.sidebar.caption("Only you can see this data; it will be cleared when you close the browser.")
personal_file = st.sidebar.file_uploader("Upload your own Excel file", type=["xlsx", "xls"], key="personal_upload")

if personal_file is not None:
    # Process only if changed
    p_file_id = f"p_{personal_file.name}_{personal_file.size}"
    if st.session_state.get("p_last_id") != p_file_id:
        st.session_state["personal_df"]  = load_data(personal_file)
        st.session_state["p_last_id"]    = p_file_id
        st.rerun()


# 2. Global Update (Admin)
st.sidebar.markdown("---")
st.sidebar.header("🛡️ Admin Server Update")
upload_pass = st.sidebar.text_input("Admin Password:", type="password")

if upload_pass == "20202024":
    admin_file = st.sidebar.file_uploader("Overwrite master file on server", type=["xlsx", "xls"], key="admin_upload")
    if admin_file is not None:
        file_id = f"a_{admin_file.name}_{admin_file.size}"
        if st.session_state.get("last_processed_id") != file_id:
            file_size_kb = round(admin_file.size / 1024, 1)
            with st.sidebar.spinner("⏳ Đang ghi đè dữ liệu máy chủ..."):
                with open(FILE_PATH, "wb") as f:
                    f.write(admin_file.getbuffer())
            
            # Xóa cache và thông tin cũ
            load_data.clear()
            if "personal_df" in st.session_state: del st.session_state["personal_df"]
            
            st.session_state["last_upload"] = {
                "name": admin_file.name,
                "size_kb": file_size_kb,
                "time": datetime.datetime.now().strftime("%H:%M:%S  %d/%m/%Y"),
            }
            st.session_state["last_processed_id"] = file_id
            st.rerun()
elif upload_pass != "":
    st.sidebar.error("❌ Sai mật khẩu!")

# Show admin success info
if "last_upload" in st.session_state:
    info = st.session_state["last_upload"]
    st.sidebar.success(f"✅ **Server Updated!**\n\n📄 {info['name']} ({info['time']})")

st.sidebar.markdown("---")

# --- Choose Data Source ---
if "personal_df" in st.session_state:
    df_raw = st.session_state["personal_df"]
    st.sidebar.info("💡 Viewing: Personal Data Mode")
    if st.sidebar.button("❌ Exit Personal Mode"):
        if "personal_df"  in st.session_state: del st.session_state["personal_df"]
        if "p_last_id"    in st.session_state: del st.session_state["p_last_id"]
        st.rerun()
else:
    df_raw = load_data(FILE_PATH)

if df_raw is not None:
    # --- Sidebar Filters ---
    st.sidebar.header(L["filters"])
    
    # Requirement: Filter specific columns D, M, Q, R, S, T->AE (including months)
    # Indices (0-based from original): 3 (D), 12 (M), 16 (Q), 17 (R), 18 (S)
    # Plus Month columns! Since we auto-pivoted, month columns are now available as "1월 금액", "2월 금액", etc.
    header_cols = df_raw.columns.tolist()
    
    # We will build filters dynamically based on original index requests and month text
    filter_options = []
    
    # Static column indices (from the requested specific list if they exist) - dimension columns only
    static_indices = [3, 12, 16]
    for i in static_indices:
        if i < len(header_cols):
            filter_options.append(header_cols[i])
            
    # Find month columns to use in the Month Selector
    # After pivot, month cols are typed like '1월 금액_1.목표예산', '1월 금액_4.비용실적'
    # We extract distinct month prefixes (e.g. '1월 금액', '2월 금액')
    import re, json
    typed_month_pat = re.compile(r'^((1[0-2]|[1-9])월 금액)_')
    seen_months = {}
    for c in header_cols:
        m = typed_month_pat.match(str(c).strip())
        if m:
            prefix = m.group(1)  # e.g. '1월 금액'
            month_num = int(re.match(r'(\d+)', prefix).group(1))
            seen_months[month_num] = prefix
    month_cols = [seen_months[k] for k in sorted(seen_months.keys())]
    
    # Clean duplicates in filter_options
    filter_options = list(dict.fromkeys(filter_options))
    
    # --- Filter Preset Save/Load ---
    PRESET_FILE = "filter_presets.json"
    if "filter_presets" not in st.session_state:
        try:
            with open(PRESET_FILE, 'r', encoding='utf-8') as f:
                st.session_state.filter_presets = json.load(f)
        except Exception:
            st.session_state.filter_presets = {}

    preset_names = list(st.session_state.filter_presets.keys())
    selected_preset = st.sidebar.selectbox(L["load_filter"], options=[""] + preset_names)
    
    if selected_preset and selected_preset in st.session_state.filter_presets:
        saved_state = st.session_state.filter_presets[selected_preset]
        saved_filter_cols = saved_state.get("filter_cols", [])
        saved_month = saved_state.get("month", None)
        saved_currency = saved_state.get("currency", "VND")
        saved_exchange_rate = saved_state.get("exchange_rate", None)
    else:
        saved_state = {}
        saved_filter_cols = filter_options[:3] if len(filter_options) >= 3 else filter_options
        saved_month = None
        saved_currency = "VND"
        saved_exchange_rate = None

    valid_saved_cols = [c for c in saved_filter_cols if c in filter_options]

    selected_filter_cols = st.sidebar.multiselect(
        L["select_cols"],
        options=filter_options,
        default=valid_saved_cols
    )
    
    # Dedicated month selector - pick which month to view
    st.sidebar.markdown("---")
    month_label_map = {prefix: f"Tháng {re.match(r'^(\d+)', prefix).group(1)} ({prefix})" for prefix in month_cols}
    month_label_options = [L["all_year"]] + list(month_label_map.values())
    
    try:
        month_idx = month_label_options.index(saved_month) if saved_month in month_label_options else 0
    except ValueError:
        month_idx = 0
        
    selected_month_label = st.sidebar.selectbox(L["view_month"], options=month_label_options, index=month_idx)
    
    # Resolve which month prefix to use
    selected_month_col = None  # None = use annual total
    for prefix, label in month_label_map.items():
        if label == selected_month_label:
            selected_month_col = prefix
            break
    
    filtered_df = df_raw.copy()
    
    # Logic for image 3 mapping
    # GT = Index 3, Dept = Index 4, Leader = Index 2, GL = Index 13, Account = Index 14
    c_gt = header_cols[3] if len(header_cols) > 3 else L["gt"]
    c_dept = header_cols[4] if len(header_cols) > 4 else L["dept"]
    c_leader = header_cols[2] if len(header_cols) > 2 else L["leader"]
    c_gl = header_cols[13] if len(header_cols) > 13 else "GL account"
    c_account = header_cols[14] if len(header_cols) > 14 else "Account"
    c_account_name = header_cols[15] if len(header_cols) > 15 else "Account Name"
    
    # Currency selection
    st.sidebar.markdown("---")
    st.sidebar.subheader(L["currency"])
    
    currency_options = {"VND": 1.0, "USD": 25000.0, "KRW": 18.0}
    
    try:
        curr_idx = list(currency_options.keys()).index(saved_currency) if saved_currency in currency_options else 0
    except ValueError:
        curr_idx = 0
    
    selected_currency = st.sidebar.selectbox(L["display_currency"], options=list(currency_options.keys()), index=curr_idx)
    
    # Allow manual rate input
    st.sidebar.caption(L["ex_rate"])
    
    default_rate = saved_exchange_rate if saved_exchange_rate is not None else currency_options[selected_currency]
    
    exchange_rate = st.sidebar.number_input(
        L["ex_rate_cap"], 
        value=float(default_rate),
        min_value=0.000001,
        format="%f"
    )
    
    curr_sym = "₫" if selected_currency == "VND" else ("₩" if selected_currency == "KRW" else "$")
    
    # Budget columns - robust matching for both Korean and English names
    def find_col(prefixes, fallback):
        for c in header_cols:
            c_str = str(c).strip()
            if any(c_str.startswith(p) for p in prefixes):
                return c
        return fallback

    c_budget_annual = find_col(["1.목표예산", "Target Budget"], "1.목표예산")
    c_increase = find_col(["2.증액예산", "Increase"], "2.증액예산")
    c_early = find_col(["3.조기 및 이월예산", "Early"], "3.조기 및 이월예산")
    c_actual_annual = find_col(["4.비용실적", "Result", "Actual"], "4.비용실적")
    
    # Override all 4 budget columns with chosen month columns if user selected a specific month
    if selected_month_col:
        # Typed month columns: e.g. "2월 금액_1.목표예산", "2월 금액_4.비용실적"
        monthly_budget_col  = f"{selected_month_col}_{c_budget_annual}"
        monthly_increase_col = f"{selected_month_col}_{c_increase}"
        monthly_early_col   = f"{selected_month_col}_{c_early}"
        monthly_actual_col  = f"{selected_month_col}_{c_actual_annual}"
        c_budget   = monthly_budget_col  if monthly_budget_col  in df_raw.columns else c_budget_annual
        c_increase = monthly_increase_col if monthly_increase_col in df_raw.columns else c_increase
        c_early    = monthly_early_col   if monthly_early_col   in df_raw.columns else c_early
        c_actual   = monthly_actual_col  if monthly_actual_col  in df_raw.columns else c_actual_annual
    else:
        c_budget = c_budget_annual
        c_actual = c_actual_annual
    
    # Collect filter values (support preset loading as defaults)
    active_filter_values = {}
    
    for col in selected_filter_cols:
        unique_vals = sorted(df_raw[col].unique().astype(str))
        saved_vals_for_col = saved_state.get("filter_values", {}).get(col, [])
        valid_saved = [v for v in saved_vals_for_col if v in unique_vals]
        selected_vals = st.sidebar.multiselect(
            f"Filter {col}:", 
            options=unique_vals,
            default=valid_saved
        )
        if selected_vals:
            active_filter_values[col] = selected_vals
            filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]
    
    # --- Save current filter preset ---
    st.sidebar.markdown("---")
    st.sidebar.subheader(L["save_preset"])
    preset_name_input = st.sidebar.text_input(L["preset_name"], placeholder="e.g. IT CHIP + Fixed")
    if st.sidebar.button(L["save_btn"]) and preset_name_input.strip():
        st.session_state.filter_presets[preset_name_input.strip()] = {
            "filter_cols": selected_filter_cols,
            "filter_values": active_filter_values,
            "month": selected_month_label,
            "currency": selected_currency,
            "exchange_rate": exchange_rate
        }
        try:
            with open(PRESET_FILE, 'w', encoding='utf-8') as f:
                import json as _json
                _json.dump(st.session_state.filter_presets, f, ensure_ascii=False, indent=2)
            st.sidebar.success(f"Saved: {preset_name_input.strip()}")
        except Exception as e:
            st.sidebar.error(f"Save failed: {e}")
            
    # Apply currency division correctly across calculation columns
    calc_cols_found = [c for c in [c_budget, c_increase, c_early, c_actual] if c in filtered_df.columns]
    for c in calc_cols_found:
        filtered_df[c] = filtered_df[c] / exchange_rate
    
    # Current GTs in the filtered view
    current_gts = filtered_df[c_gt].unique().tolist()
    
    # --- Metrics ---
    # Recalculate based on found column names
    avail_cols = filtered_df.columns.tolist()
    if all(c in avail_cols for c in [c_budget, c_increase, c_early, c_actual]):
        total_target = filtered_df[c_budget].sum()
        total_actual = filtered_df[c_actual].sum()
        total_all_budget = (filtered_df[c_budget] + filtered_df[c_increase] + filtered_df[c_early]).sum()
        usage_pct = (total_actual / total_all_budget * 100) if total_all_budget > 0 else 0
        
        m1, m2, m3 = st.columns(3)
        if curr_sym in ["₫", "₩"]:
            m1.metric(f"{L['total_budget']} ({selected_currency})", f"{curr_sym} {total_all_budget:,.0f}")
            m2.metric(f"{L['total_actual']}", f"{curr_sym} {total_actual:,.0f}")
        else:
            m1.metric(f"{L['total_budget']} ({selected_currency})", f"{curr_sym}{total_all_budget:,.2f}")
            m2.metric(f"{L['total_actual']}", f"{curr_sym}{total_actual:,.2f}")
        m3.metric(L["usage_pct"], f"{usage_pct:.1f}%")
    else:
        # Fallback calculation if some columns missing but actual/budget exist
        total_target = filtered_df[c_budget].sum() if c_budget in avail_cols else 0
        total_actual = filtered_df[c_actual].sum() if c_actual in avail_cols else 0
        total_all_budget = total_target # Simple fallback
        usage_pct = (total_actual / total_all_budget * 100) if total_all_budget > 0 else 0
        
        st.warning(f"Note: Some budget columns not found. Using available data.")
        m1, m2, m3 = st.columns(3)
        if curr_sym in ["₫", "₩"]:
            m1.metric(L["total_budget"], f"{curr_sym} {total_target:,.0f}")
            m2.metric(L["total_actual"], f"{curr_sym} {total_actual:,.0f}")
        else:
            m1.metric(L["total_budget"], f"{curr_sym}{total_target:,.2f}")
            m2.metric(L["total_actual"], f"{curr_sym}{total_actual:,.2f}")
        m3.metric(L["usage_pct"], f"{usage_pct:.1f}%")

    st.markdown("---")
    
    # --- Visualizations ---
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader(L["chart_overview"])
        if all(c in filtered_df.columns for c in [c_budget, c_actual]):
            num_items = filtered_df[c_dept].nunique() if c_dept in filtered_df.columns else 0
            
            if num_items > 15 or num_items == 0:
                summary_data = pd.DataFrame({
                    "Category": [L["target"], L["result"]],
                    "Amount": [total_target, total_actual]
                })
                fig = px.bar(
                    summary_data, x="Category", y="Amount", color="Category",
                    color_discrete_map={L["target"]: "#1e293b", L["result"]: "#3b82f6"},
                    template="plotly_white"
                )
            else:
                chart_data = filtered_df.groupby(c_dept)[[c_budget, c_actual]].sum().reset_index()
                fig = px.bar(
                    chart_data, x=c_dept, y=[c_budget, c_actual],
                    labels={"value": "Amount", "variable": "Category"},
                    barmode="group", color_discrete_sequence=["#1e293b", "#3b82f6"],
                    template="plotly_white"
                )
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader(L["chart_ratio"])
        if 'usage_pct' in locals():
            fig_pie = px.pie(
                names=["Used", "Remaining"],
                values=[total_actual, max(0, total_all_budget - total_actual)],
                hole=0.5,
                color_discrete_sequence=["#ef4444" if usage_pct > 100 else "#3b82f6", "#e2e8f0"]
            )
            st.plotly_chart(fig_pie, use_container_width=True)



    # --- GT Summary Breakdown (When multiple or all GTs are selected) ---
    if len(current_gts) != 1:
        st.markdown("---")
        st.subheader("GT Group Details")
        
        sum_cols = [c for c in [c_budget, c_increase, c_early, c_actual] if c in filtered_df.columns]
        if sum_cols:
            group_cols = [c_gt, c_leader]
            valid_group_cols = [c for c in group_cols if c in filtered_df.columns]

            gt_data = filtered_df.groupby(valid_group_cols)[sum_cols].sum().reset_index()

            
            rename_ops = {
                c_gt: L["gt"],
                c_leader: L["leader"],
                c_budget: L["target"],
                c_increase: L["increase"],
                c_early: L["early"],
                c_actual: L["result"]
            }
            gt_display = gt_data.rename(columns=rename_ops)
            
            c_res = L["result"]
            c_tar = L["target"]
            if c_res in gt_display.columns and c_tar in gt_display.columns:
                gt_display[L["usage_tar"]] = (gt_display[c_res] / gt_display[c_tar]).replace([float('inf'), -float('inf')], 0).fillna(0)
            else:
                gt_display[L["usage_tar"]] = 0
                
            vol_cols = [c for c in [L["target"], L["increase"], L["early"]] if c in gt_display.columns]
            if c_res in gt_display.columns and vol_cols:
                gt_tot = sum(gt_display[c] for c in vol_cols)
                if isinstance(gt_tot, pd.Series):
                    gt_display[L["usage_all"]] = (gt_display[c_res] / gt_tot).replace([float('inf'), -float('inf')], 0).fillna(0)
                else: 
                    gt_display[L["usage_all"]] = gt_display[c_res] / gt_tot if gt_tot != 0 else 0
            else:
                gt_display[L["usage_all"]] = 0
                
            final_cols = [L["gt"], L["leader"], L["target"], L["increase"], L["early"], L["result"], L["usage_tar"], L["usage_all"]]
            gt_display = gt_display[[c for c in final_cols if c in gt_display.columns]]
            
            st.dataframe(get_styled_df(gt_display, curr_sym), use_container_width=True, hide_index=True)

    # --- GL Account Breakdown (Conditional) ---
    # Triggered if only one GT is present in the current filtered view
    elif len(current_gts) == 1:
        st.markdown("---")
        st.subheader(f"{L['gl_detail']}: {current_gts[0]}")
        
        if c_gl in filtered_df.columns:
            # Group by GL account (safely handling missing budget columns)
            sum_cols = [c for c in [c_budget, c_increase, c_early, c_actual] if c in filtered_df.columns]
            
            if not sum_cols:
                gl_data = pd.DataFrame(columns=[c_gl])
            else:
                # Group by GL, GT, Dept, Leader to retain full information for all departments
                group_cols = [c_gl, c_gt, c_dept, c_leader]
                valid_group_cols = [c for c in group_cols if c in filtered_df.columns]

                gl_data = filtered_df.groupby(valid_group_cols)[sum_cols].sum().reset_index()
                
            # Rename columns back carefully safely
            rename_ops = {
                c_gl: L["gl"],
                c_gt: L["gt"],
                c_dept: L["dept"],
                c_leader: L["leader"],
                c_budget: L["target"],
                c_increase: L["increase"],
                c_early: L["early"],
                c_actual: L["result"]
            }
            gl_display = gl_data.rename(columns=rename_ops)
            gl_display = translate_content(gl_display)

            
            # Recalculate percentages safely using renamed columns
            c_res = L["result"]
            c_tar = L["target"]
            if c_res in gl_display.columns and c_tar in gl_display.columns:
                gl_display[L["usage_tar"]] = (gl_display[c_res] / gl_display[c_tar]).replace([float('inf'), -float('inf')], 0).fillna(0)
            else:
                gl_display[L["usage_tar"]] = 0
                
            vol_cols = [c for c in [L["target"], L["increase"], L["early"]] if c in gl_display.columns]
            if c_res in gl_display.columns and vol_cols:
                gl_tot = sum(gl_display[c] for c in vol_cols)
                if isinstance(gl_tot, pd.Series):
                    gl_display[L["usage_all"]] = (gl_display[c_res] / gl_tot).replace([float('inf'), -float('inf')], 0).fillna(0)
                else: 
                    gl_display[L["usage_all"]] = gl_display[c_res] / gl_tot if gl_tot != 0 else 0
            else:
                gl_display[L["usage_all"]] = 0
            
            # Ensure correct order
            final_cols = [L["gt"], L["dept"], L["leader"], L["gl"], L["target"], L["increase"], L["early"], L["result"], L["usage_tar"], L["usage_all"]]
            gl_display = gl_display[[c for c in final_cols if c in gl_display.columns]]
            
            st.dataframe(get_styled_df(gl_display, curr_sym), use_container_width=True, hide_index=True)
            
            # --- Over-budget Drill-down ---
            if L["usage_tar"] in gl_display.columns and L["gl"] in gl_display.columns:
                over_budget_rows = gl_display[gl_display[L["usage_tar"]] > 1.0]
                if not over_budget_rows.empty:
                    st.markdown("---")
                    st.subheader(L["over_budget_title"])
                    st.caption(L["over_budget_cap"])
                    
                    over_gls = over_budget_rows[L["gl"]].tolist()
                    
                    if st.session_state.lang == "Tiếng Việt":
                        rev_gl_map = {v: k for k, v in GL_VALUE_MAP.items()}
                        over_gls_kor = [rev_gl_map.get(g, g) for g in over_gls]
                    else:
                        over_gls_kor = over_gls
                        
                    drill_df = filtered_df[filtered_df[c_gl].isin(over_gls_kor)]
                    
                    if c_account in drill_df.columns:
                        if c_account in drill_df.columns:
                            drill_group_cols = [c for c in [c_gl, c_account, c_account_name, c_gt, c_dept, c_leader] if c in drill_df.columns]
                            drill_data = drill_df.groupby(drill_group_cols)[sum_cols].sum().reset_index()
                            
                            drill_display = drill_data.rename(columns=rename_ops)
                            
                            # Rename account ID and Name
                            drill_display.rename(columns={
                                c_account: L["acc_id"],
                                c_account_name: L["acc_name"]
                            }, inplace=True)
                            
                            drill_display = translate_content(drill_display)

                            
                            # Recalculate percentages for drill down
                            if c_res in drill_display.columns and c_tar in drill_display.columns:
                                drill_display[L["usage_tar"]] = (drill_display[c_res] / drill_display[c_tar]).replace([float('inf'), -float('inf')], 0).fillna(0)
                            else:
                                drill_display[L["usage_tar"]] = 0
                                
                            # Usage All
                            if c_res in drill_display.columns and vol_cols:
                                tot_d = sum(drill_display[c] for c in vol_cols)
                                if isinstance(tot_d, pd.Series):
                                    drill_display[L["usage_all"]] = (drill_display[c_res] / tot_d).replace([float('inf'), -float('inf')], 0).fillna(0)
                                else:
                                    drill_display[L["usage_all"]] = drill_display[c_res] / tot_d if tot_d != 0 else 0
                            else:
                                drill_display[L["usage_all"]] = 0
                            
                            # Order columns correctly
                            avail_drill_cols = drill_display.columns.tolist()
                            final_drill_cols = [L["gt"], L["dept"], L["leader"], L["gl"]]
                            if L["acc_id"] in avail_drill_cols: final_drill_cols.append(L["acc_id"])
                            if L["acc_name"] in avail_drill_cols: final_drill_cols.append(L["acc_name"])
                            final_drill_cols.extend([L["target"], L["increase"], L["early"], L["result"], L["usage_tar"], L["usage_all"]])
                            
                            drill_display = drill_display[[c for c in final_drill_cols if c in drill_display.columns]]
                            
                            # Show ALL detail accounts in the over-budget group (don't filter by usage)
                            if L["result"] in drill_display.columns:
                                drill_display = drill_display[drill_display[L["result"]] != 0]
                            
                        if not drill_display.empty:
                            st.dataframe(get_styled_df(drill_display, curr_sym), use_container_width=True, hide_index=True)
                        else:
                            st.info("No spending found in this account group for the selected period.")
                    else:
                        st.warning("Detail Account column (Column 15/O) not found for drill-down.")
        else:
            st.warning("Column 'GL account' not found.")
 


else:
    st.error(f"File '{FILE_PATH}' not found. Please ensure the Excel file exists in the directory.")

st.sidebar.markdown("---")
st.sidebar.info(f"📍 Machine IP: 172.16.122.37\nAccess via LAN: http://172.16.122.37:8501")
