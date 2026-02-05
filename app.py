import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
import re
from datetime import timedelta
from io import BytesIO

def _theme():
    return {
        "config": {
            "background": "rgba(0,0,0,0)",
            "view": {"stroke": "transparent"},
            "axis": {
                "labelColor": "#3b3b35",
                "titleColor": "#3b3b35",
                "gridColor": "#e6dfd3",
                "labelFont": "Noto Sans TC",
                "titleFont": "Noto Sans TC",
            },
            "title": {
                "color": "#1d1b16",
                "font": "Noto Serif TC",
                "fontSize": 18,
            },
            "range": {
                "category": [
                    "#2b7a78",
                    "#f2b134",
                    "#8f5e3b",
                    "#9bc1bc",
                    "#d95d39",
                    "#5f6c37",
                    "#7a4e2d",
                ]
            },
        }
    }

alt.themes.register("custom_theme", _theme)
alt.themes.enable("custom_theme")

st.set_page_config(page_title="顧客關係經營分析", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@600;700&family=Noto+Sans+TC:wght@400;500;600&display=swap');
:root {
  --ink: #1d1b16;
  --muted: #6b6b63;
  --accent: #2b7a78;
  --accent-2: #f2b134;
  --card: #ffffff;
  --line: #e6dfd3;
  --shadow: 0 6px 18px rgba(33, 28, 20, 0.08);
}
html, body, [class*="css"] {
  font-family: "Noto Sans TC", "PingFang TC", "Microsoft JhengHei", sans-serif;
  color: var(--ink);
}
.stApp {
  background:
    radial-gradient(1200px 600px at 10% -10%, #f6efe3 0%, transparent 60%),
    radial-gradient(1000px 500px at 110% 10%, #eaf3f2 0%, transparent 55%),
    linear-gradient(180deg, #f7f4ee 0%, #f3f0ea 35%, #f9f7f3 100%);
}
h1, h2, h3, h4 {
  font-family: "Noto Serif TC", "Noto Sans TC", serif;
  letter-spacing: 0.4px;
}
h1 {
  font-size: 2.2rem;
  color: var(--ink);
}
div[data-testid="stSidebar"] {
  background: #faf7f2;
  border-right: 1px solid var(--line);
}
div[data-testid="stMetric"] {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 10px 12px;
  box-shadow: var(--shadow);
}
div[data-testid="stMetric"] label {
  color: var(--muted);
  font-weight: 600;
}
.section-note {
  padding: 8px 12px;
  background: rgba(43, 122, 120, 0.08);
  border: 1px solid rgba(43, 122, 120, 0.2);
  border-radius: 12px;
  color: #1f4f4e;
  font-size: 0.95rem;
}
.stCaption {
  color: var(--muted);
}
.stDataFrame, .stTable {
  border: 1px solid var(--line);
  border-radius: 12px;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("顧客關係經營分析")
st.markdown('<div class="section-note">圖表優先，表格放在最後。指標固定：新客流失率、熟客經營力、新客經營力、空窗率、合作穩定度。</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("輸入")
    member_file = st.file_uploader("上傳 會員名單.xlsx（選填）", type=["xlsx"], key="member")
    bill_files = st.file_uploader("上傳 帳單紀錄.xlsx（可多檔）", type=["xlsx"], key="bill", accept_multiple_files=True)
    st.divider()
    include_types = st.multiselect(
        "計算包含結帳類型",
        ["服務", "票券", "商品", "儲值金"],
        default=["服務", "票券"],
    )
    chart_top_n = st.number_input("圖表顯示前 N 名（0=全部）", min_value=0, max_value=100, value=0, step=1)
    store_chart_type = st.selectbox("分店比較圖表", ["群組直條圖", "熱度圖", "堆疊條圖"])

st.write("""
本工具會：
- 以全品牌帳單歷史找出新客，並計算 60 天內是否回店（同分店）
- 依分店/師傅呈現新客流失率、熟客經營力、新客經營力、空窗率與合作穩定度
- 提供圖表與排行榜，快速看出差異
回指判定：T2=30 天、T3=60 天（固定）
空窗率計算：依項目分鐘估算時長（1～30=0.5；31～60=1；61～90=1.5，以此類推），月上限 168 小時
""")

if not bill_files:
    st.info("請先上傳帳單檔。")
    st.stop()

CHURN_DAYS = 60
T2_DAYS = 30
T3_DAYS = 60

@st.cache_data(show_spinner=False)
def load_member(file):
    return pd.read_excel(file, sheet_name="會員名單")

@st.cache_data(show_spinner=False)
def load_bill(file, sheets):
    xls = pd.ExcelFile(file)
    available = [s for s in sheets if s in xls.sheet_names]
    frames = []
    for s in available:
        frames.append(pd.read_excel(file, sheet_name=s))
    if not frames:
        return pd.DataFrame(), available
    return pd.concat(frames, ignore_index=True), available

def infer_store_name(filename):
    stem = Path(filename).stem
    name = stem
    name = re.sub(r"帳單紀錄", "", name)
    name = re.sub(r"帳單|紀錄", "", name)
    name = re.sub(r"\d{4}-\d{2}-\d{2}", "", name)
    name = re.sub(r"\d{8,}", "", name)
    name = re.sub(r"[_\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else stem

def load_bills(files, sheets):
    frames = []
    used = set()
    for f in files:
        df, used_sheets = load_bill(f, sheets)
        if not df.empty:
            store_name = infer_store_name(getattr(f, "name", "未命名分店"))
            df["來源檔案"] = getattr(f, "name", "未命名檔案")
            if "分店" not in df.columns:
                df["分店"] = store_name
            else:
                df["分店"] = df["分店"].fillna(store_name)
            frames.append(df)
        used.update(used_sheets)
    if not frames:
        return pd.DataFrame(), list(used)
    return pd.concat(frames, ignore_index=True), list(used)

bills, used_sheets = load_bills(bill_files, include_types)

if bills.empty:
    st.error("帳單檔中找不到選擇的工作表。")
    st.stop()

# Normalize phone numbers

def norm_digits(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)):
        if np.isnan(x):
            return None
        s = str(int(x)) if float(x).is_integer() else str(x)
    else:
        s = str(x)
    digits = re.sub(r"\D+", "", s)
    return digits or None

def extract_minutes(item_text):
    if pd.isna(item_text):
        return 0
    mins = re.findall(r"(\d+)\s*分鐘", str(item_text))
    if not mins:
        return 0
    return sum(int(m) for m in mins)

def minutes_to_hours(mins):
    if mins <= 0:
        return 0.0
    # 1~30=0.5, 31~60=1, 61~90=1.5, ...
    units = int((mins - 1) // 30 + 1)
    return units * 0.5

def render_bar_chart(df, category_col, value_col, title, color="#2b7a78", top_n=0, value_format="percent", orient="vertical", ascending=False):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return 0, 0
    chart_df = df[[category_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=ascending)
    if top_n and len(chart_df) > top_n:
        chart_df = chart_df.head(int(top_n))
    chart_df[value_col] = chart_df[value_col].astype(float)
    height = 320 if orient == "vertical" else min(520, 28 * len(chart_df) + 40)
    if value_format == "percent":
        axis_format = "%"
        label_format = ".1%"
        tooltip = [category_col, alt.Tooltip(value_col, format=".2%")]
    elif value_format == "number1":
        axis_format = ",.1f"
        label_format = ".1f"
        tooltip = [category_col, alt.Tooltip(value_col, format=",.1f")]
    else:
        axis_format = ",.0f"
        label_format = ".0f"
        tooltip = [category_col, alt.Tooltip(value_col, format=",.0f")]

    if orient == "vertical":
        base = alt.Chart(chart_df).encode(
            x=alt.X(
                f"{category_col}:N",
                sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
                title="",
                axis=alt.Axis(labelAngle=-30),
            ),
            y=alt.Y(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="center", dy=-6, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    else:
        base = alt.Chart(chart_df).encode(
            y=alt.Y(
                f"{category_col}:N",
                sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
                title="",
            ),
            x=alt.X(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="left", dx=4, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)
    return len(chart_df), len(df)

def render_rank_bar(df, name_col, value_col, title, ascending, value_format, color, top_n=6):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return
    chart_df = df[[name_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=ascending).head(top_n)
    chart_df[value_col] = chart_df[value_col].astype(float)

    if value_format == "percent":
        axis_format = "%"
        label_format = ".1%"
        tooltip = [name_col, alt.Tooltip(value_col, format=".2%")]
    elif value_format == "number1":
        axis_format = ",.1f"
        label_format = ".1f"
        tooltip = [name_col, alt.Tooltip(value_col, format=",.1f")]
    else:
        axis_format = ",.0f"
        label_format = ".0f"
        tooltip = [name_col, alt.Tooltip(value_col, format=",.0f")]

    base = alt.Chart(chart_df).encode(
        y=alt.Y(
            f"{name_col}:N",
            sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
            title="",
        ),
        x=alt.X(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
        tooltip=tooltip,
    )
    bars = base.mark_bar(color=color)
    labels = base.mark_text(align="left", dx=4, color="#333").encode(
        text=alt.Text(f"{value_col}:Q", format=label_format)
    )
    height = min(420, 28 * len(chart_df) + 40)
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)


for df, col in [(bills, "國碼"), (bills, "電話號碼")]:
    if col in df.columns:
        bills[col] = bills[col].apply(norm_digits)

if "國碼" not in bills.columns or "電話號碼" not in bills.columns:
    st.error("帳單檔缺少 '國碼' 或 '電話號碼' 欄位。")
    st.stop()

bills["phone_key"] = bills["國碼"].fillna("") + "-" + bills["電話號碼"].fillna("")

valid_bills = bills[bills["phone_key"].str.contains("-") & (bills["phone_key"] != "-")].copy()

valid_bills["結帳操作時間"] = pd.to_datetime(valid_bills["結帳操作時間"], errors="coerce")

# Merge member data (optional)
merged = valid_bills.copy()
if member_file:
    members = load_member(member_file)
    for df, col in [(members, "國碼"), (members, "手機號碼")]:
        if col in df.columns:
            df[col] = df[col].apply(norm_digits)
    if "國碼" not in members.columns or "手機號碼" not in members.columns:
        st.error("會員名單缺少 '國碼' 或 '手機號碼' 欄位。")
        st.stop()
    members["phone_key"] = members["國碼"].fillna("") + "-" + members["手機號碼"].fillna("")
    valid_members = members[members["phone_key"].str.contains("-") & (members["phone_key"] != "-")].copy()
    member_cols = ["phone_key", "來店次數", "會員姓名"]
    member_cols = [c for c in member_cols if c in valid_members.columns]
    merged = valid_bills.merge(valid_members[member_cols], on="phone_key", how="left")

# First checkout per phone
merged_sorted = merged.sort_values("結帳操作時間")
first_checkout = merged_sorted.groupby("phone_key", as_index=False).first()

# New customer definition (全品牌首次結帳)
new_first = first_checkout.copy()

end_date_all = merged["結帳操作時間"].max()

# Choose member name column
name_col = "會員姓名"
if "會員姓名" not in new_first.columns:
    # fallback if merge suffixes exist
    if "會員姓名_y" in new_first.columns:
        new_first[name_col] = new_first["會員姓名_y"]
    elif "會員姓名_x" in new_first.columns:
        new_first[name_col] = new_first["會員姓名_x"]

# Designer column
if "設計師" not in new_first.columns:
    st.error("帳單檔缺少 '設計師' 欄位，無法分師傅。")
    st.stop()

# Store column (optional)
has_store = "分店" in new_first.columns
if not has_store:
    st.warning("帳單檔缺少 '分店' 欄位，將無法依分店分組。")

matured = new_first[new_first["matured"]].copy()

# Sidebar filters
with st.sidebar:
    st.header("篩選")
    if has_store:
        store_options = sorted([s for s in merged["分店"].dropna().unique()])
        store_filter = st.multiselect("分店", store_options, default=store_options)
    else:
        store_filter = None
    if has_store and store_filter is not None:
        base_for_designers = merged[merged["分店"].isin(store_filter)]
    else:
        base_for_designers = merged
    designer_options = sorted([s for s in base_for_designers["設計師"].dropna().unique()])
    exclude_designers = st.multiselect("排除師傅", designer_options, default=[])
    include_options = [d for d in designer_options if d not in set(exclude_designers)]
    designer_filter = st.multiselect("師傅", include_options, default=include_options)

merged_store = merged.copy()
if has_store and store_filter is not None:
    merged_store = merged_store[merged_store["分店"].isin(store_filter)]
merged_store = merged_store.sort_values("結帳操作時間")
merged_designer = merged_store[merged_store["設計師"].isin(designer_filter)].copy()

# 新客以全品牌判定，請確保已上傳全品牌資料
st.caption("新客口徑為全品牌歷史首次結帳；若未上傳全品牌帳單，可能高估新客與流失率。")

end_date = merged_store["結帳操作時間"].max()
if pd.isna(end_date):
    st.error("篩選後沒有可用資料。")
    st.stop()

# 新客（全品牌首購）→ 同分店回店
new_first_store = new_first.copy()
if has_store and store_filter is not None:
    new_first_store = new_first_store[new_first_store["分店"].isin(store_filter)]

checkouts_by_phone_store = merged_store.groupby(["phone_key", "分店"])["結帳操作時間"].apply(list)
churn_flags = []
return_days_store = []
for _, row in new_first_store.iterrows():
    pk = row["phone_key"]
    store = row.get("分店")
    first_time = row["結帳操作時間"]
    times = checkouts_by_phone_store.get((pk, store), [])
    next_time = None
    for t in times:
        if pd.isna(t):
            continue
        if t > first_time:
            next_time = t
            break
    if next_time is not None:
        return_days_store.append((next_time - first_time).days)
        churn_flags.append(next_time > first_time + timedelta(days=CHURN_DAYS))
    else:
        return_days_store.append(np.nan)
        churn_flags.append(True)
new_first_store["return_days_store"] = return_days_store
new_first_store["churn"] = churn_flags
new_first_store["matured"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=CHURN_DAYS) <= end_date

filtered_new_first = new_first_store[new_first_store["設計師"].isin(designer_filter)].copy()

def add_repeat_flags(df, list_map, key_cols, first_col, t2, t3):
    days2 = []
    days3 = []
    repeat2 = []
    repeat3 = []
    for _, row in df.iterrows():
        key = tuple(row[c] for c in key_cols)
        first_time = row[first_col]
        times = list_map.get(key, [])
        times = [t for t in times if pd.notna(t)]
        times.sort()
        after = [t for t in times if t > first_time]
        d2 = (after[0] - first_time).days if len(after) >= 1 else np.nan
        d3 = (after[1] - first_time).days if len(after) >= 2 else np.nan
        days2.append(d2)
        days3.append(d3)
        repeat2.append(d2 <= t2)
        repeat3.append(d3 <= t3)
    df["days_to_2nd"] = days2
    df["days_to_3rd"] = days3
    df["repeat2"] = repeat2
    df["repeat3"] = repeat3
    return df

checkouts_by_phone_store_designer = (
    merged_store.groupby(["phone_key", "分店", "設計師"])["結帳操作時間"]
    .apply(list)
)

new_first_store["matured_2"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=T2_DAYS) <= end_date
new_first_store["matured_3"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=T3_DAYS) <= end_date
new_first_store = add_repeat_flags(
    new_first_store,
    checkouts_by_phone_store_designer,
    ["phone_key", "分店", "設計師"],
    "結帳操作時間",
    T2_DAYS,
    T3_DAYS,
)

end_ts = pd.to_datetime(end_date)
start_ts_3m = end_ts - pd.DateOffset(months=3)
start_ts_6m = end_ts - pd.DateOffset(months=6)

# 新客經營力（近 3 個月）
new_recent = new_first_store[new_first_store["結帳操作時間"] >= start_ts_3m].copy()
new_recent_churn = new_recent[new_recent["matured"]].copy()
new_churn_by_designer = (
    new_recent_churn.groupby("設計師")
    .agg(
        new_customers_3m=("phone_key", "count"),
        new_churned_3m=("churn", "sum"),
    )
    .reset_index()
)
new_churn_by_designer["new_churn_rate_3m"] = np.where(
    new_churn_by_designer["new_customers_3m"] > 0,
    new_churn_by_designer["new_churned_3m"] / new_churn_by_designer["new_customers_3m"],
    np.nan,
)

new_recent2 = new_recent[new_recent["matured_2"]].copy()
new_by_designer = (
    new_recent2.groupby("設計師")
    .agg(
        new_repeat_base_3m=("phone_key", "count"),
        new_repeat2_count=("repeat2", "sum"),
    )
    .reset_index()
)
new_by_designer["new_repeat_rate_3m"] = np.where(
    new_by_designer["new_repeat_base_3m"] > 0,
    new_by_designer["new_repeat2_count"] / new_by_designer["new_repeat_base_3m"],
    np.nan,
)

new_recent3 = new_recent[new_recent["matured_3"]].copy()
new_deep = (
    new_recent3.groupby("設計師")
    .agg(new_deep_rate_3m=("repeat3", "mean"), new_deep_n=("repeat3", "count"))
    .reset_index()
)

# 熟客經營力（近 3 個月，同分店熟客）
history = merged_store[merged_store["結帳操作時間"] < start_ts_3m]
familiar_keys = (
    history.groupby(["phone_key", "分店"])
    .size()
    .reset_index(name="visits_before")
)
familiar_keys = familiar_keys[familiar_keys["visits_before"] >= 2][["phone_key", "分店"]]

if familiar_keys.empty:
    familiar_first = pd.DataFrame(columns=["phone_key", "分店", "設計師", "baseline_time"])
else:
    familiar_visits = merged_store[merged_store["結帳操作時間"] >= start_ts_3m]
    familiar_visits = familiar_visits.merge(familiar_keys, on=["phone_key", "分店"], how="inner")
    familiar_first = (
        familiar_visits.sort_values("結帳操作時間")
        .groupby(["phone_key", "分店", "設計師"], as_index=False)
        .first()
        .rename(columns={"結帳操作時間": "baseline_time"})
    )

if not familiar_first.empty:
    familiar_first["matured_2"] = familiar_first["baseline_time"] + pd.Timedelta(days=T2_DAYS) <= end_date
    familiar_first["matured_3"] = familiar_first["baseline_time"] + pd.Timedelta(days=T3_DAYS) <= end_date
    familiar_first = add_repeat_flags(
        familiar_first,
        checkouts_by_phone_store_designer,
        ["phone_key", "分店", "設計師"],
        "baseline_time",
        T2_DAYS,
        T3_DAYS,
    )

if familiar_first.empty:
    fam_recent2 = familiar_first.copy()
    fam_recent3 = familiar_first.copy()
else:
    fam_recent2 = familiar_first[familiar_first["matured_2"]].copy()
    fam_recent3 = familiar_first[familiar_first["matured_3"]].copy()
fam_by_designer = (
    fam_recent2.groupby("設計師")
    .agg(
        familiar_customers_3m=("phone_key", "count"),
        familiar_repeat2_count=("repeat2", "sum"),
    )
    .reset_index()
)
fam_by_designer["familiar_repeat_rate_3m"] = np.where(
    fam_by_designer["familiar_customers_3m"] > 0,
    fam_by_designer["familiar_repeat2_count"] / fam_by_designer["familiar_customers_3m"],
    np.nan,
)

fam_deep = (
    fam_recent3.groupby("設計師")
    .agg(familiar_deep_rate_3m=("repeat3", "mean"), familiar_deep_n=("repeat3", "count"))
    .reset_index()
)

# 合併師傅指標
designer_metrics = (
    new_churn_by_designer
    .merge(new_by_designer, on="設計師", how="outer")
    .merge(new_deep, on="設計師", how="outer")
    .merge(fam_by_designer, on="設計師", how="outer")
    .merge(fam_deep, on="設計師", how="outer")
)

# Store monthly summary (average per month)
store_monthly_avg = None
if has_store:
    store_month = new_first_store[new_first_store["matured"]].copy()
    store_month["month"] = store_month["結帳操作時間"].dt.to_period("M").astype(str)
    month_summary = (
        store_month.groupby(["分店", "month"])
        .agg(matured_new_customers=("phone_key", "count"), churned=("churn", "sum"))
        .reset_index()
    )
    month_summary["retained"] = month_summary["matured_new_customers"] - month_summary["churned"]
    month_summary["repeat_rate"] = np.where(
        month_summary["matured_new_customers"] > 0,
        month_summary["retained"] / month_summary["matured_new_customers"],
        np.nan,
    )
    store_monthly_avg = (
        month_summary.groupby("分店")
        .agg(
            月平均新客數=("matured_new_customers", "mean"),
            月平均流失數=("churned", "mean"),
            月平均留住數=("retained", "mean"),
            平均回店率=("repeat_rate", "mean"),
            月份數=("month", "nunique"),
        )
        .reset_index()
    )

summary_by_store = None
summary_by_store_designer = None
if has_store:
    matured_store = new_first_store[new_first_store["matured"]].copy()
    summary_by_store = (
        matured_store.groupby("分店")
        .agg(
            matured_new_customers=("phone_key", "count"),
            churned=("churn", "sum"),
        )
        .reset_index()
    )
    summary_by_store["churn_rate"] = summary_by_store["churned"] / summary_by_store["matured_new_customers"]
    summary_by_store["repeat_rate"] = 1 - summary_by_store["churn_rate"]

    summary_by_store_designer = (
        matured_store[matured_store["設計師"].isin(designer_filter)]
        .groupby(["分店", "設計師"])
        .agg(
            matured_new_customers=("phone_key", "count"),
            churned=("churn", "sum"),
        )
        .reset_index()
    )
    summary_by_store_designer["churn_rate"] = (
        summary_by_store_designer["churned"] / summary_by_store_designer["matured_new_customers"]
    )
    summary_by_store_designer["repeat_rate"] = 1 - summary_by_store_designer["churn_rate"]

# Vacancy metrics (monthly, 168h cap)
vacancy_monthly = None
vacancy_recent = None
if "項目" in merged_store.columns:
    time_df = merged_store.copy()
    time_df["duration_minutes"] = time_df["項目"].apply(extract_minutes)
    time_df["duration_hours"] = time_df["duration_minutes"].apply(minutes_to_hours)
    time_df["month"] = time_df["結帳操作時間"].dt.to_period("M").astype(str)
    if has_store:
        group_cols = ["分店", "設計師", "month"]
    else:
        group_cols = ["設計師", "month"]
    vacancy_monthly = (
        time_df.groupby(group_cols)["duration_hours"]
        .sum()
        .reset_index()
    )
    vacancy_monthly["vacancy_rate"] = (1 - vacancy_monthly["duration_hours"] / 168.0).clip(lower=0, upper=1)
    vm = vacancy_monthly.copy()
    vm["month_start"] = pd.to_datetime(vm["month"] + "-01")
    vm = vm[vm["month_start"] >= start_ts_3m]
    vacancy_recent = (
        vm.groupby("設計師")["vacancy_rate"]
        .mean()
        .reset_index()
        .rename(columns={"vacancy_rate": "vacancy_rate_3m"})
    )
    designer_metrics = designer_metrics.merge(vacancy_recent, on="設計師", how="left")

# 合作穩定度（近 6 個月）
stability_by_designer = None
stability_df = merged_store.copy()
stability_df["month"] = stability_df["結帳操作時間"].dt.to_period("M").astype(str)
stability_df["date"] = stability_df["結帳操作時間"].dt.date
active_days = (
    stability_df.groupby(["設計師", "month"])["date"]
    .nunique()
    .reset_index(name="active_days")
)
active_days["month_start"] = pd.to_datetime(active_days["month"] + "-01")
active_days_6m = active_days[active_days["month_start"] >= start_ts_6m]

stability_by_designer = (
    active_days_6m.groupby("設計師")["active_days"]
    .agg(active_days_avg_6m="mean", active_days_cv_6m=lambda s: s.std() / s.mean() if s.mean() else np.nan)
    .reset_index()
)

if vacancy_monthly is not None:
    hours_df = vacancy_monthly.copy()
    hours_df = hours_df.rename(columns={"duration_hours": "service_hours"})
    hours_df["month_start"] = pd.to_datetime(hours_df["month"] + "-01")
    hours_6m = hours_df[hours_df["month_start"] >= start_ts_6m]
    hours_summary = (
        hours_6m.groupby("設計師")["service_hours"]
        .agg(service_hours_avg_6m="mean", service_hours_cv_6m=lambda s: s.std() / s.mean() if s.mean() else np.nan)
        .reset_index()
    )
    stability_by_designer = stability_by_designer.merge(hours_summary, on="設計師", how="left")

last_tx = (
    merged_store.groupby("設計師")["結帳操作時間"]
    .max()
    .reset_index()
    .rename(columns={"結帳操作時間": "last_tx"})
)
last_tx["days_since_last_tx"] = (end_ts - last_tx["last_tx"]).dt.days
stability_by_designer = stability_by_designer.merge(last_tx, on="設計師", how="left")
designer_metrics = designer_metrics.merge(stability_by_designer, on="設計師", how="left")
designer_metrics_filtered = designer_metrics[designer_metrics["設計師"].isin(designer_filter)].copy()

overall = {
    "matured_new_customers": int(len(new_first_store[new_first_store["matured"]])),
    "churned_matured": int(new_first_store.loc[new_first_store["matured"], "churn"].sum()),
}
overall["retained_matured"] = overall["matured_new_customers"] - overall["churned_matured"]
overall["churn_rate_matured"] = (
    overall["churned_matured"] / overall["matured_new_customers"]
    if overall["matured_new_customers"]
    else None
)
overall["repeat_rate_matured"] = (
    1 - overall["churn_rate_matured"]
    if overall["churn_rate_matured"] is not None
    else None
)

st.subheader("總覽摘要")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("滿60天新客數", overall["matured_new_customers"])
col2.metric("流失人數", overall["churned_matured"])
col3.metric("留住人數", overall["retained_matured"])
col4.metric("流失率", f"{overall['churn_rate_matured']:.2%}" if overall["churn_rate_matured"] is not None else "-")
col5.metric("回店率(同分店)", f"{overall['repeat_rate_matured']:.2%}" if overall["repeat_rate_matured"] is not None else "-")

st.caption(f"資料截止時間：{end_date}")

if has_store and store_monthly_avg is not None and not store_monthly_avg.empty:
    st.subheader("各分店摘要（每月平均）")
    st.caption("僅計入滿 60 天的新客 cohort（最嚴謹）。")
    for _, row in store_monthly_avg.iterrows():
        store_name = row["分店"]
        avg_new = row["月平均新客數"]
        avg_churn = row["月平均流失數"]
        avg_retained = row["月平均留住數"]
        avg_repeat = row["平均回店率"]
        repeat_text = f"{avg_repeat:.2%}" if pd.notna(avg_repeat) else "-"
        st.write(
            f"- {store_name}：每月平均新客 {avg_new:.1f} 人、流失 {avg_churn:.1f} 人、留住 {avg_retained:.1f} 人、回店率 {repeat_text}"
        )

    st.subheader("分店指標對比（每月平均）")
    metric_base = store_monthly_avg.copy()
    if chart_top_n and chart_top_n > 0:
        top_stores = (
            metric_base.sort_values("月平均新客數", ascending=False)
            .head(int(chart_top_n))["分店"]
            .tolist()
        )
        metric_base = metric_base[metric_base["分店"].isin(top_stores)]

    metric_long = metric_base.melt(
        id_vars=["分店"],
        value_vars=["月平均新客數", "月平均流失數", "月平均留住數"],
        var_name="指標",
        value_name="人數",
    )
    metric_long["指標"] = metric_long["指標"].replace(
        {
            "月平均新客數": "新客(人)",
            "月平均流失數": "流失(人)",
            "月平均留住數": "留住(人)",
        }
    )

    if store_chart_type == "群組直條圖":
        group_chart = alt.Chart(metric_long).mark_bar().encode(
            x=alt.X("分店:N", sort="-y", title="", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("人數:Q", title="人數(每月平均)"),
            color=alt.Color("指標:N", title="指標"),
            xOffset="指標:N",
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        st.altair_chart(group_chart, use_container_width=True)
    elif store_chart_type == "熱度圖":
        heat = alt.Chart(metric_long).mark_rect().encode(
            x=alt.X("指標:N", title=""),
            y=alt.Y("分店:N", title=""),
            color=alt.Color("人數:Q", title="人數(每月平均)"),
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        text = alt.Chart(metric_long).mark_text(color="white").encode(
            x="指標:N",
            y="分店:N",
            text=alt.Text("人數:Q", format=".1f"),
        )
        st.altair_chart(heat + text, use_container_width=True)
    else:
        stack_long = metric_long[metric_long["指標"].isin(["流失(人)", "留住(人)"])].copy()
        stack = alt.Chart(stack_long).mark_bar().encode(
            x=alt.X("分店:N", sort="-y", title="", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("人數:Q", stack=True, title="每月平均新客(人)"),
            color=alt.Color("指標:N", title="構成"),
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        totals = store_monthly_avg.copy()
        totals["總人"] = totals["月平均流失數"] + totals["月平均留住數"]
        total_text = alt.Chart(totals).mark_text(dy=-6, color="#333").encode(
            x=alt.X("分店:N", sort="-y"),
            y=alt.Y("總人:Q"),
            text=alt.Text("總人:Q", format=".1f"),
        )
        st.altair_chart(stack + total_text, use_container_width=True)

    if chart_top_n and chart_top_n > 0:
        st.caption(f"分店比較圖僅顯示前 {int(chart_top_n)} 名（依每月平均新客數排序）。")

    # 移除 Small Multiples 依需求

st.subheader("師傅指標比較（最近 3 個月）")
metric_df = designer_metrics_filtered.copy()
service_cv = metric_df["service_hours_cv_6m"] if "service_hours_cv_6m" in metric_df.columns else pd.Series(np.nan, index=metric_df.index)
active_cv = metric_df["active_days_cv_6m"] if "active_days_cv_6m" in metric_df.columns else pd.Series(np.nan, index=metric_df.index)
metric_df["合作穩定度(CV)"] = np.where(service_cv.notna(), service_cv, active_cv)

metric_options = {
    "新客流失率(60天，低越好)": ("new_churn_rate_3m", "percent", True),
    "新客回指率(30天，高越好)": ("new_repeat_rate_3m", "percent", False),
    "新客深度回指率(60天，高越好)": ("new_deep_rate_3m", "percent", False),
    "熟客回指率(30天，高越好)": ("familiar_repeat_rate_3m", "percent", False),
    "熟客深度回指率(60天，高越好)": ("familiar_deep_rate_3m", "percent", False),
    "空窗率(3M，低越好)": ("vacancy_rate_3m", "percent", True),
    "合作穩定度(CV，低越好)": ("合作穩定度(CV)", "number1", True),
    "最近有單距今(天，低越好)": ("days_since_last_tx", "number1", True),
}

metric_choice = st.selectbox("選擇指標", list(metric_options.keys()))
metric_col, metric_fmt, metric_asc = metric_options[metric_choice]
metric_view = metric_df[["設計師", metric_col]].dropna().copy()
render_bar_chart(
    metric_view,
    "設計師",
    metric_col,
    metric_choice,
    color="#4e79a7",
    top_n=chart_top_n,
    value_format=metric_fmt,
    orient="horizontal",
    ascending=metric_asc,
)

st.subheader("個別師傅狀態（最近 3 個月）")
st.caption("新客＝全品牌首次；回店口徑＝同分店；回指 T2=30 天、T3=60 天。")

if not designer_filter:
    st.info("目前沒有可顯示的師傅。")
else:
    designer_select = st.selectbox("選擇師傅", designer_filter)
    selected_row = designer_metrics_filtered[designer_metrics_filtered["設計師"] == designer_select].copy()

    if selected_row.empty:
        st.info("此師傅在最近 3 個月內沒有足夠資料。")
    else:
        r = selected_row.iloc[0]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("新客流失率(60天)", f"{r['new_churn_rate_3m']:.2%}" if pd.notna(r.get("new_churn_rate_3m")) else "-")
        c2.metric("新客回指率(30天)", f"{r['new_repeat_rate_3m']:.2%}" if pd.notna(r.get("new_repeat_rate_3m")) else "-")
        c3.metric("熟客回指率(30天)", f"{r['familiar_repeat_rate_3m']:.2%}" if pd.notna(r.get("familiar_repeat_rate_3m")) else "-")
        c4.metric("空窗率(3M)", f"{r['vacancy_rate_3m']:.2%}" if pd.notna(r.get("vacancy_rate_3m")) else "-")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("新客深度回指(60天)", f"{r['new_deep_rate_3m']:.2%}" if pd.notna(r.get("new_deep_rate_3m")) else "-")
        c2.metric("熟客深度回指(60天)", f"{r['familiar_deep_rate_3m']:.2%}" if pd.notna(r.get("familiar_deep_rate_3m")) else "-")
        c3.metric("最近有單距今(天)", f"{int(r['days_since_last_tx'])}" if pd.notna(r.get("days_since_last_tx")) else "-")
        if pd.notna(r.get("service_hours_cv_6m")):
            c4.metric("合作穩定度(工時CV)", f"{r['service_hours_cv_6m']:.2f}")
        else:
            c4.metric("合作穩定度(出勤CV)", f"{r['active_days_cv_6m']:.2f}" if pd.notna(r.get("active_days_cv_6m")) else "-")

        st.caption("回指率以「同師傅、同分店」計算；合作穩定度 CV 越低代表越穩定。")

        # 師傅比較：熟客 vs 新客回指率
        compare_df = designer_metrics_filtered.dropna(subset=["new_repeat_rate_3m", "familiar_repeat_rate_3m"]).copy()
        if not compare_df.empty:
            base = alt.Chart(compare_df).mark_circle(size=60, color="#9aa0a6").encode(
                x=alt.X("familiar_repeat_rate_3m:Q", axis=alt.Axis(format="%", title="熟客回指率(30天)")),
                y=alt.Y("new_repeat_rate_3m:Q", axis=alt.Axis(format="%", title="新客回指率(30天)")),
                tooltip=["設計師", alt.Tooltip("familiar_repeat_rate_3m:Q", format=".2%"), alt.Tooltip("new_repeat_rate_3m:Q", format=".2%")],
            )
            highlight = alt.Chart(compare_df[compare_df["設計師"] == designer_select]).mark_circle(size=140, color="#e15759").encode(
                x="familiar_repeat_rate_3m:Q",
                y="new_repeat_rate_3m:Q",
                tooltip=["設計師"],
            )
            st.altair_chart(base + highlight, use_container_width=True)

        # 成長/下滑曲線（新客/熟客回指率）
        st.subheader("師傅成長/下滑曲線（回指率）")
        new_cohort = new_first_store[new_first_store["設計師"] == designer_select].copy()
        new_cohort["cohort_month"] = new_cohort["結帳操作時間"].dt.to_period("M").dt.to_timestamp()
        new_cohort2 = new_cohort[new_cohort["matured_2"]].copy()
        new_trend = (
            new_cohort2.groupby("cohort_month")
            .agg(rate=("repeat2", "mean"), n=("repeat2", "count"))
            .reset_index()
        )
        new_trend["指標"] = "新客回指率(30天)"

        fam_cohort = familiar_first[familiar_first["設計師"] == designer_select].copy()
        if not fam_cohort.empty:
            fam_cohort["cohort_month"] = fam_cohort["baseline_time"].dt.to_period("M").dt.to_timestamp()
            fam_cohort2 = fam_cohort[fam_cohort["matured_2"]].copy()
            fam_trend = (
                fam_cohort2.groupby("cohort_month")
                .agg(rate=("repeat2", "mean"), n=("repeat2", "count"))
                .reset_index()
            )
            fam_trend["指標"] = "熟客回指率(30天)"
            trend_long = pd.concat([new_trend, fam_trend], ignore_index=True)
        else:
            trend_long = new_trend.copy()

        if not trend_long.empty:
            line = alt.Chart(trend_long.dropna(subset=["rate"])).mark_line(point=True).encode(
                x=alt.X("cohort_month:T", title="首次來訪月份"),
                y=alt.Y("rate:Q", axis=alt.Axis(format="%"), title="回指率"),
                color=alt.Color("指標:N", title=""),
                tooltip=["cohort_month:T", "指標", alt.Tooltip("rate:Q", format=".2%")],
            ).properties(height=280)
            st.altair_chart(line, use_container_width=True)
        else:
            st.info("目前沒有足夠的回指資料繪製趨勢。")

st.subheader("師傅排行榜（Top 6，最近 3 個月）")
st.caption("新客以全品牌首購；回店口徑同分店；回指 T2=30/T3=60。")
col_good, col_watch = st.columns(2)
has_hours_cv = "service_hours_cv_6m" in designer_metrics_filtered.columns and designer_metrics_filtered["service_hours_cv_6m"].notna().any()

with col_good:
    st.markdown("<span style='color:#2ca02c;font-weight:700;'>表現較佳</span>", unsafe_allow_html=True)
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_churn_rate_3m"]),
        "設計師",
        "new_churn_rate_3m",
        "新客流失率最低(60天)",
        ascending=True,
        value_format="percent",
        color="#2ca02c",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["familiar_repeat_rate_3m"]),
        "設計師",
        "familiar_repeat_rate_3m",
        "熟客回指率最高(30天)",
        ascending=False,
        value_format="percent",
        color="#2ca02c",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_repeat_rate_3m"]),
        "設計師",
        "new_repeat_rate_3m",
        "新客回指率最高(30天)",
        ascending=False,
        value_format="percent",
        color="#2ca02c",
    )
    if vacancy_recent is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最低(3M)",
            ascending=True,
            value_format="percent",
            color="#2ca02c",
        )
    if has_hours_cv:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["service_hours_cv_6m"]),
            "設計師",
            "service_hours_cv_6m",
            "合作穩定度最佳(工時CV最低)",
            ascending=True,
            value_format="number1",
            color="#2ca02c",
        )
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["active_days_cv_6m"]),
            "設計師",
            "active_days_cv_6m",
            "合作穩定度最佳(出勤CV最低)",
            ascending=True,
            value_format="number1",
            color="#2ca02c",
        )

with col_watch:
    st.markdown("<span style='color:#d62728;font-weight:700;'>需關注</span>", unsafe_allow_html=True)
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_churn_rate_3m"]),
        "設計師",
        "new_churn_rate_3m",
        "新客流失率最高(60天)",
        ascending=False,
        value_format="percent",
        color="#d62728",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["familiar_repeat_rate_3m"]),
        "設計師",
        "familiar_repeat_rate_3m",
        "熟客回指率最低(30天)",
        ascending=True,
        value_format="percent",
        color="#d62728",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_repeat_rate_3m"]),
        "設計師",
        "new_repeat_rate_3m",
        "新客回指率最低(30天)",
        ascending=True,
        value_format="percent",
        color="#d62728",
    )
    if vacancy_recent is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最高(3M)",
            ascending=False,
            value_format="percent",
            color="#d62728",
        )
    if has_hours_cv:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["service_hours_cv_6m"]),
            "設計師",
            "service_hours_cv_6m",
            "合作穩定度較弱(工時CV最高)",
            ascending=False,
            value_format="number1",
            color="#d62728",
        )
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["active_days_cv_6m"]),
            "設計師",
            "active_days_cv_6m",
            "合作穩定度較弱(出勤CV最高)",
            ascending=False,
            value_format="number1",
            color="#d62728",
        )

st.subheader("回店時間分布（同分店）")
ret_series = filtered_new_first["return_days_store"].dropna() if "return_days_store" in filtered_new_first.columns else pd.Series([], dtype=float)
if ret_series.empty:
    st.info("目前沒有可計算的回訪資料。")
else:
    p50 = ret_series.quantile(0.5)
    p70 = ret_series.quantile(0.7)
    p80 = ret_series.quantile(0.8)
    p90 = ret_series.quantile(0.9)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P50(天)", f"{p50:.0f}")
    c2.metric("P70(天)", f"{p70:.0f}")
    c3.metric("P80(天)", f"{p80:.0f}")
    c4.metric("P90(天)", f"{p90:.0f}")

    hist_df = pd.DataFrame({"回店天數": ret_series})
    hist = alt.Chart(hist_df).mark_bar(color="#4e79a7").encode(
        x=alt.X("回店天數:Q", bin=alt.Bin(step=7), title="回店天數(天)"),
        y=alt.Y("count():Q", title="人數"),
        tooltip=[alt.Tooltip("count()", title="人數")],
    ).properties(height=260)

    p_df = pd.DataFrame({"x": [p70, p80], "label": ["P70", "P80"]})
    rules = alt.Chart(p_df).mark_rule(color="#f28e2b").encode(x="x:Q")
    labels = alt.Chart(p_df).mark_text(align="left", dx=4, dy=-4, color="#f28e2b").encode(
        x="x:Q", y=alt.value(0), text="label:N"
    )
    st.altair_chart(hist + rules + labels, use_container_width=True)

st.subheader("詳細表格")
st.caption("依照你的篩選條件，以下是完整明細表格。")
store_table = None
store_designer_table = None
display_vacancy = None

# 師傅彙總（近3月）
st.markdown("**師傅彙總（近 3 月）**")
if designer_metrics_filtered.empty:
    st.info("目前沒有足夠的近 3 月資料。")
    designer_table = None
else:
    designer_table = designer_metrics_filtered.copy()
    service_cv = designer_table["service_hours_cv_6m"] if "service_hours_cv_6m" in designer_table.columns else pd.Series(np.nan, index=designer_table.index)
    active_cv = designer_table["active_days_cv_6m"] if "active_days_cv_6m" in designer_table.columns else pd.Series(np.nan, index=designer_table.index)
    designer_table["合作穩定度(CV)"] = np.where(service_cv.notna(), service_cv, active_cv)
    designer_table = designer_table.rename(
        columns={
            "設計師": "師傅",
            "new_customers_3m": "新客數(3M,滿60天)",
            "new_churn_rate_3m": "新客流失率(60天)",
            "new_repeat_rate_3m": "新客回指率(30天)",
            "new_deep_rate_3m": "新客深度回指率(60天)",
            "familiar_customers_3m": "熟客數(3M,滿30天)",
            "familiar_repeat_rate_3m": "熟客回指率(30天)",
            "familiar_deep_rate_3m": "熟客深度回指率(60天)",
            "vacancy_rate_3m": "空窗率(3M)",
            "days_since_last_tx": "最近有單距今(天)",
        }
    )
    cols = [
        "師傅",
        "新客數(3M,滿60天)",
        "新客流失率(60天)",
        "新客回指率(30天)",
        "新客深度回指率(60天)",
        "熟客數(3M,滿30天)",
        "熟客回指率(30天)",
        "熟客深度回指率(60天)",
        "空窗率(3M)",
        "合作穩定度(CV)",
        "最近有單距今(天)",
    ]
    cols = [c for c in cols if c in designer_table.columns]
    st.dataframe(designer_table[cols].sort_values("新客流失率(60天)", ascending=True), use_container_width=True)
    st.caption("合作穩定度：若缺少項目分鐘，將以出勤 CV 代替工時 CV。")

# 分店彙總
if has_store:
    st.markdown("**分店彙總**")
    store_table = summary_by_store.rename(
        columns={
            "分店": "分店",
            "matured_new_customers": "滿60天新客數",
            "churned": "流失數",
            "churn_rate": "流失率",
            "repeat_rate": "回店率",
        }
    )
    st.dataframe(store_table.sort_values("流失率", ascending=False), use_container_width=True)

    st.markdown("**分店 x 師傅彙總**")
    store_designer_table = summary_by_store_designer.rename(
        columns={
            "分店": "分店",
            "設計師": "師傅",
            "matured_new_customers": "滿60天新客數",
            "churned": "流失數",
            "churn_rate": "流失率",
            "repeat_rate": "回店率",
        }
    )
    st.dataframe(store_designer_table.sort_values("流失率", ascending=False), use_container_width=True)

# 空窗率（月，168 小時上限）
st.markdown("**空窗率（月，168 小時上限）**")
if vacancy_monthly is None:
    st.warning("帳單檔缺少 '項目' 欄位，無法估算服務時數與空窗率。")
else:
    display_vacancy = vacancy_monthly.copy()
    if has_store and store_filter is not None:
        display_vacancy = display_vacancy[display_vacancy["分店"].isin(store_filter)]
    display_vacancy = display_vacancy[display_vacancy["設計師"].isin(designer_filter)]
    display_vacancy = display_vacancy.rename(
        columns={
            "分店": "分店",
            "設計師": "師傅",
            "month": "月份",
            "duration_hours": "服務時數",
            "vacancy_rate": "空窗率",
        }
    )
    st.dataframe(display_vacancy.sort_values(["月份", "師傅"]), use_container_width=True)

# 流失名單
st.markdown("**流失名單**")
show_churned_only = st.checkbox("只看流失者", value=True)

if show_churned_only:
    detail = filtered_new_first[filtered_new_first["churn"]].copy()
else:
    detail = filtered_new_first.copy()

if has_store and store_filter is not None:
    detail = detail[detail["分店"].isin(store_filter)]
detail = detail[detail["設計師"].isin(designer_filter)]

detail_display = detail.copy()
detail_display["是否滿期"] = detail_display["matured"].map({True: "是", False: "否"})
detail_display["是否流失"] = detail_display["churn"].map({True: "是", False: "否"})
detail_display["電話"] = detail_display["phone_key"]
detail_display = detail_display.rename(columns={"設計師": "師傅", "結帳操作時間": "首單時間"})

display_cols = ["電話", name_col, "分店", "師傅", "首單時間", "是否滿期", "是否流失"]
display_cols = [c for c in display_cols if c in detail_display.columns]
st.dataframe(detail_display[display_cols].sort_values("首單時間"), use_container_width=True)

# Download Excel
st.subheader("下載報表")

output = BytesIO()
overall_df = pd.DataFrame([{
    "滿60天新客數": overall["matured_new_customers"],
    "流失人數": overall["churned_matured"],
    "留住人數": overall["retained_matured"],
    "流失率": overall["churn_rate_matured"],
    "回店率": overall["repeat_rate_matured"],
}])
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    overall_df.to_excel(writer, index=False, sheet_name="總覽")
    if designer_table is not None:
        designer_table.to_excel(writer, index=False, sheet_name="師傅彙總(3M)")
    if store_table is not None:
        store_table.to_excel(writer, index=False, sheet_name="分店彙總")
    if store_designer_table is not None:
        store_designer_table.to_excel(writer, index=False, sheet_name="分店師傅彙總")
    detail_display[display_cols].to_excel(writer, index=False, sheet_name="流失名單")
    if display_vacancy is not None:
        display_vacancy.to_excel(writer, index=False, sheet_name="空窗率(月)")

st.download_button(
    label="下載 Excel",
    data=output.getvalue(),
    file_name="customer_relationship_analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
