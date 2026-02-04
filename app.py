import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
import re
from datetime import timedelta
from io import BytesIO

st.set_page_config(page_title="新客流失率分析", layout="wide")

st.title("顧客關係經營分析")

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
    churn_days = st.number_input("流失判定天數", min_value=1, max_value=365, value=60, step=1)
    new_customer_mode = st.selectbox(
        "新客判定方式",
        ["帳單歷史首次結帳（建議）", "會員來店次數 = 1"],
    )
    chart_top_n = st.number_input("圖表顯示前 N 名（0=全部）", min_value=0, max_value=100, value=0, step=1)
    store_chart_type = st.selectbox("分店比較圖表", ["群組直條圖", "熱度圖", "堆疊條圖"])

st.write("""
本工具會：
- 以帳單歷史找出新客，並計算 60 天內是否回訪
- 依分店/師傅呈現流失率、回訪率與空窗率
- 提供圖表與排行榜，快速看出差異
空窗率計算：依項目分鐘估算時長（1～30=0.5；31～60=1；61～90=1.5，以此類推），月上限 168 小時
""")

if not bill_files:
    st.info("請先上傳帳單檔。")
    st.stop()

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

def render_bar_chart(df, category_col, value_col, title, color="#4e79a7", top_n=0, value_format="percent", orient="vertical"):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return 0, 0
    chart_df = df[[category_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=False)
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
            x=alt.X(f"{category_col}:N", sort="-y", title="", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="center", dy=-6, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    else:
        base = alt.Chart(chart_df).encode(
            y=alt.Y(f"{category_col}:N", sort="-x", title=""),
            x=alt.X(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="left", dx=4, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)
    return len(chart_df), len(df)

def render_rank_table(df, name_col, value_col, title, ascending, value_fmt, top_n=3, text_color=None):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return
    rank_df = df[[name_col, value_col]].dropna().copy()
    rank_df = rank_df.sort_values(value_col, ascending=ascending).head(top_n)
    rank_df = rank_df.reset_index(drop=True)
    if value_fmt == "percent":
        formatter = lambda v: f"{v:.2%}"
    elif value_fmt == "int":
        formatter = lambda v: f"{int(v)}"
    elif value_fmt == "number1":
        formatter = lambda v: f"{v:.1f}"
    else:
        formatter = lambda v: f"{v}"
    st.markdown(f"**{title}**")
    lines = []
    for idx, row in rank_df.iterrows():
        line = f"{idx+1}. {row[name_col]} — {formatter(row[value_col])}"
        if text_color:
            line = f"<span style='color:{text_color}'>{line}</span>"
        lines.append(line)
    st.markdown("<br>".join(lines), unsafe_allow_html=True)

def compute_return_days(new_first_df, checkout_lists):
    days = []
    for _, row in new_first_df.iterrows():
        pk = row["phone_key"]
        first_time = row["結帳操作時間"]
        times = checkout_lists.get(pk, [])
        next_time = None
        for t in times:
            if pd.isna(t):
                continue
            if t > first_time:
                next_time = t
                break
        if next_time is not None:
            days.append((next_time - first_time).days)
    return pd.Series(days, name="return_days")

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

# New customer definition
if new_customer_mode == "會員來店次數 = 1":
    if not member_file:
        st.error("使用「會員來店次數 = 1」需要上傳會員名單。")
        st.stop()
    if "來店次數" not in first_checkout.columns:
        st.error("會員名單缺少 '來店次數' 欄位。")
        st.stop()
    new_first = first_checkout[first_checkout["來店次數"] == 1].copy()
else:
    new_first = first_checkout.copy()

# Build list of all checkouts per phone
checkouts = merged_sorted.groupby("phone_key")["結帳操作時間"].apply(list)

# Churn
cutoff_days = int(churn_days)
churn_flags = []
for _, row in new_first.iterrows():
    pk = row["phone_key"]
    first_time = row["結帳操作時間"]
    times = checkouts.get(pk, [])
    cutoff = first_time + timedelta(days=cutoff_days)
    has_return = False
    for t in times:
        if pd.isna(t):
            continue
        if t > first_time and t <= cutoff:
            has_return = True
            break
    churn_flags.append(not has_return)
new_first["churn"] = churn_flags

# Maturity
end_date = merged["結帳操作時間"].max()
new_first["matured"] = new_first["結帳操作時間"] + pd.Timedelta(days=cutoff_days) <= end_date

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
    designer_options = sorted([s for s in merged["設計師"].dropna().unique()])
    exclude_designers = st.multiselect("排除師傅", designer_options, default=[])
    include_options = [d for d in designer_options if d not in set(exclude_designers)]
    designer_filter = st.multiselect("師傅", include_options, default=include_options)

filtered_new_first = new_first.copy()
if has_store and store_filter is not None:
    filtered_new_first = filtered_new_first[filtered_new_first["分店"].isin(store_filter)]
filtered_new_first = filtered_new_first[filtered_new_first["設計師"].isin(designer_filter)]

filtered_matured = filtered_new_first[filtered_new_first["matured"]].copy()

# 師傅回指 cohort：同一位師傅的首次來訪（new-to-therapist）
tx_filtered = merged.copy()
if has_store and store_filter is not None:
    tx_filtered = tx_filtered[tx_filtered["分店"].isin(store_filter)]
tx_filtered = tx_filtered[tx_filtered["設計師"].isin(designer_filter)]
tx_filtered = tx_filtered.sort_values("結帳操作時間")

therapist_groups = tx_filtered.groupby(["phone_key", "設計師"], as_index=False)
first_visits = therapist_groups.first()
visit_lists = therapist_groups["結帳操作時間"].apply(list).reset_index(name="visit_times")
therapist_first = first_visits.merge(visit_lists, on=["phone_key", "設計師"], how="left")
therapist_first = therapist_first.rename(columns={"結帳操作時間": "first_time"})

def days_to_n(times, n):
    if not isinstance(times, list) or len(times) < n:
        return np.nan
    return (times[n - 1] - times[0]).days

therapist_first["days_to_2nd"] = therapist_first["visit_times"].apply(lambda t: days_to_n(t, 2))
therapist_first["days_to_3rd"] = therapist_first["visit_times"].apply(lambda t: days_to_n(t, 3))

summary_by_designer = (
    filtered_matured.groupby("設計師")
    .agg(
        matured_new_customers=("phone_key", "count"),
        churned=("churn", "sum"),
    )
    .reset_index()
)
summary_by_designer["churn_rate"] = summary_by_designer["churned"] / summary_by_designer["matured_new_customers"]
summary_by_designer["repeat_rate"] = 1 - summary_by_designer["churn_rate"]

if has_store:
    summary_by_store = (
        filtered_matured.groupby("分店")
        .agg(
            matured_new_customers=("phone_key", "count"),
            churned=("churn", "sum"),
        )
        .reset_index()
    )
    summary_by_store["churn_rate"] = summary_by_store["churned"] / summary_by_store["matured_new_customers"]
    summary_by_store["repeat_rate"] = 1 - summary_by_store["churn_rate"]

    summary_by_store_designer = (
        filtered_matured.groupby(["分店", "設計師"])
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

# Designer-level metrics for rankings (僅計入滿 60 天新客)
designer_metrics = summary_by_designer.copy()
designer_metrics["new_customers"] = designer_metrics["matured_new_customers"]

# Store monthly summary (average per month)
store_monthly_avg = None
if has_store:
    store_month = filtered_matured.copy()
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
            平均回訪率=("repeat_rate", "mean"),
            月份數=("month", "nunique"),
        )
        .reset_index()
    )

# Vacancy metrics (monthly, 168h cap)
vacancy_monthly = None
vacancy_by_designer = None
if "項目" in merged.columns:
    time_df = merged.copy()
    time_df["duration_minutes"] = time_df["項目"].apply(extract_minutes)
    time_df["duration_hours"] = time_df["duration_minutes"].apply(minutes_to_hours)
    time_df["month"] = time_df["結帳操作時間"].dt.to_period("M").astype(str)
    if has_store and store_filter is not None:
        time_df = time_df[time_df["分店"].isin(store_filter)]
    time_df = time_df[time_df["設計師"].isin(designer_filter)]
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
    vacancy_by_designer = (
        vacancy_monthly.groupby("設計師")["vacancy_rate"]
        .mean()
        .reset_index()
    )
    designer_metrics = designer_metrics.merge(vacancy_by_designer, on="設計師", how="left")

overall = {
    "matured_new_customers": int(len(filtered_matured)),
    "churned_matured": int(filtered_matured["churn"].sum()),
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
col5.metric("回訪率", f"{overall['repeat_rate_matured']:.2%}" if overall["repeat_rate_matured"] is not None else "-")

st.caption(f"資料截止時間：{end_date}")

if has_store and store_monthly_avg is not None and not store_monthly_avg.empty:
    st.subheader("各分店摘要（每月平均）")
    st.caption("僅計入滿 60 天的新客 cohort（最嚴謹）。")
    for _, row in store_monthly_avg.iterrows():
        store_name = row["分店"]
        avg_new = row["月平均新客數"]
        avg_churn = row["月平均流失數"]
        avg_retained = row["月平均留住數"]
        avg_repeat = row["平均回訪率"]
        repeat_text = f"{avg_repeat:.2%}" if pd.notna(avg_repeat) else "-"
        st.write(
            f"- {store_name}：每月平均新客 {avg_new:.1f} 人、流失 {avg_churn:.1f} 人、留住 {avg_retained:.1f} 人、回訪率 {repeat_text}"
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

st.subheader("個別師傅綜合狀態（最近 3 個月）")
st.caption("回指口徑：同一位師傅。僅計入滿 T 天的新客 cohort。")

# 回指時間分布（第2次 / 第3次） → 用來決定 T2/T3
dist2 = therapist_first["days_to_2nd"].dropna()
dist3 = therapist_first["days_to_3rd"].dropna()

if dist2.empty:
    st.info("目前沒有可計算的第2次回指資料。")
    t2_default = 60
else:
    p50_2 = dist2.quantile(0.5)
    p70_2 = dist2.quantile(0.7)
    p80_2 = dist2.quantile(0.8)
    p90_2 = dist2.quantile(0.9)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("第2次 P50(天)", f"{p50_2:.0f}")
    c2.metric("第2次 P70(天)", f"{p70_2:.0f}")
    c3.metric("第2次 P80(天)", f"{p80_2:.0f}")
    c4.metric("第2次 P90(天)", f"{p90_2:.0f}")
    t2_default = int(round(p70_2))

if dist3.empty:
    st.info("目前沒有可計算的第3次回指資料。")
    t3_default = 120
else:
    p50_3 = dist3.quantile(0.5)
    p70_3 = dist3.quantile(0.7)
    p80_3 = dist3.quantile(0.8)
    p90_3 = dist3.quantile(0.9)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("第3次 P50(天)", f"{p50_3:.0f}")
    c2.metric("第3次 P70(天)", f"{p70_3:.0f}")
    c3.metric("第3次 P80(天)", f"{p80_3:.0f}")
    c4.metric("第3次 P90(天)", f"{p90_3:.0f}")
    t3_default = int(round(p70_3))

t2_days = st.number_input("第2次回指判定天數（T2）", min_value=1, max_value=365, value=int(t2_default), step=1)
t3_days = st.number_input("第3次回指判定天數（T3）", min_value=1, max_value=365, value=int(t3_default), step=1)

therapist_first["matured_2"] = therapist_first["first_time"] + pd.Timedelta(days=int(t2_days)) <= end_date
therapist_first["matured_3"] = therapist_first["first_time"] + pd.Timedelta(days=int(t3_days)) <= end_date
therapist_first["repeat2"] = therapist_first["days_to_2nd"] <= int(t2_days)
therapist_first["repeat3"] = therapist_first["days_to_3rd"] <= int(t3_days)

# 最近三個月的師傅指標（回指口徑）
end_ts = pd.to_datetime(end_date)
start_ts = end_ts - pd.DateOffset(months=3)
recent = therapist_first[therapist_first["first_time"] >= start_ts].copy()

recent2 = recent[recent["matured_2"]].copy()
recent3 = recent[recent["matured_3"]].copy()

recent_by_designer = (
    recent2.groupby("設計師")
    .agg(
        new_customers_3m=("phone_key", "count"),
        repeat2_count=("repeat2", "sum"),
    )
    .reset_index()
)
recent_by_designer["repeat_rate_3m"] = np.where(
    recent_by_designer["new_customers_3m"] > 0,
    recent_by_designer["repeat2_count"] / recent_by_designer["new_customers_3m"],
    np.nan,
)

deep_repeat = (
    recent3.groupby("設計師")
    .agg(deep_repeat_rate_3m=("repeat3", "mean"), deep_n=("repeat3", "count"))
    .reset_index()
)
recent_by_designer = recent_by_designer.merge(deep_repeat, on="設計師", how="left")

# 空窗率（最近三個月平均）
vacancy_recent = None
if vacancy_monthly is not None and not vacancy_monthly.empty:
    vm = vacancy_monthly.copy()
    vm["month_start"] = pd.to_datetime(vm["month"] + "-01")
    vm = vm[vm["month_start"] >= start_ts]
    vacancy_recent = (
        vm.groupby("設計師")["vacancy_rate"]
        .mean()
        .reset_index()
        .rename(columns={"vacancy_rate": "vacancy_rate_3m"})
    )
    recent_by_designer = recent_by_designer.merge(vacancy_recent, on="設計師", how="left")

# 綜合分數（百分位加權）
score_df = recent_by_designer.copy()
if not score_df.empty:
    score_df["repeat_pct"] = score_df["repeat_rate_3m"].rank(pct=True)
    score_df["new_pct"] = score_df["new_customers_3m"].rank(pct=True)
    if "vacancy_rate_3m" in score_df.columns:
        score_df["vacancy_pct"] = 1 - score_df["vacancy_rate_3m"].rank(pct=True)
    else:
        score_df["vacancy_pct"] = np.nan

    def calc_score(row):
        weights = {"repeat_pct": 0.4, "new_pct": 0.35, "vacancy_pct": 0.25}
        vals = {k: row[k] for k in weights if pd.notna(row[k])}
        if not vals:
            return np.nan
        wsum = sum(weights[k] for k in vals.keys())
        return sum(vals[k] * weights[k] for k in vals.keys()) / wsum * 100

    score_df["score"] = score_df.apply(calc_score, axis=1)

designer_select = st.selectbox("選擇師傅", designer_filter if designer_filter else [])
mode = st.radio("模式", ["綜合分數", "四象限圖"], horizontal=True)

selected_row = score_df[score_df["設計師"] == designer_select].copy()
if selected_row.empty:
    st.info("此師傅在最近 3 個月內沒有滿 T2 的新客資料。")
else:
    r = selected_row.iloc[0]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("新客數(3M)", int(r["new_customers_3m"]))
    c2.metric("回指率(第2次)", f"{r['repeat_rate_3m']:.2%}" if pd.notna(r["repeat_rate_3m"]) else "-")
    c3.metric("深度回指率(第3次)", f"{r['deep_repeat_rate_3m']:.2%}" if pd.notna(r.get("deep_repeat_rate_3m")) else "-")
    c4.metric("空窗率(3M)", f"{r['vacancy_rate_3m']:.2%}" if "vacancy_rate_3m" in r and pd.notna(r["vacancy_rate_3m"]) else "-")
    c5.metric("綜合分數", f"{r['score']:.1f}" if pd.notna(r["score"]) else "-")

    if mode == "綜合分數":
        st.caption("綜合分數（升階參考）＝回指率(第2次)40% + 新客數35% + 空窗率25%（百分位加權）")
    else:
        plot_df = score_df.dropna(subset=["repeat_rate_3m", "new_customers_3m"]).copy()
        if not plot_df.empty:
            x_med = plot_df["repeat_rate_3m"].median()
            y_med = plot_df["new_customers_3m"].median()
            base = alt.Chart(plot_df).mark_circle(size=60, color="#9aa0a6").encode(
                x=alt.X("repeat_rate_3m:Q", axis=alt.Axis(format="%", title="回指率(第2次)")),
                y=alt.Y("new_customers_3m:Q", title="新客數(3M)"),
                tooltip=["設計師", alt.Tooltip("repeat_rate_3m:Q", format=".2%"), "new_customers_3m:Q"],
            )
            highlight = alt.Chart(plot_df[plot_df["設計師"] == designer_select]).mark_circle(size=140, color="#e15759").encode(
                x="repeat_rate_3m:Q",
                y="new_customers_3m:Q",
                tooltip=["設計師"],
            )
            vline = alt.Chart(pd.DataFrame({"x": [x_med]})).mark_rule(color="#666").encode(x="x:Q")
            hline = alt.Chart(pd.DataFrame({"y": [y_med]})).mark_rule(color="#666").encode(y="y:Q")
            st.altair_chart(base + vline + hline + highlight, use_container_width=True)

    # 成長/下滑曲線（cohort 月）
    st.subheader("師傅成長/下滑曲線（回指率）")
    cohort = therapist_first[therapist_first["設計師"] == designer_select].copy()
    cohort["cohort_month"] = cohort["first_time"].dt.to_period("M").dt.to_timestamp()
    cohort2 = cohort[cohort["matured_2"]].copy()
    cohort3 = cohort[cohort["matured_3"]].copy()

    m2 = (
        cohort2.groupby("cohort_month")
        .agg(rate2=("repeat2", "mean"), n2=("repeat2", "count"))
        .reset_index()
    )
    m3 = (
        cohort3.groupby("cohort_month")
        .agg(rate3=("repeat3", "mean"), n3=("repeat3", "count"))
        .reset_index()
    )

    trend = pd.merge(m2, m3, on="cohort_month", how="outer").sort_values("cohort_month")
    trend_long = trend.melt(
        id_vars=["cohort_month"],
        value_vars=["rate2", "rate3"],
        var_name="指標",
        value_name="回指率",
    )
    trend_long["指標"] = trend_long["指標"].replace(
        {"rate2": "回指率(第2次)", "rate3": "深度回指率(第3次)"}
    )

    if trend_long["回指率"].notna().any():
        line = alt.Chart(trend_long.dropna(subset=["回指率"])).mark_line(point=True).encode(
            x=alt.X("cohort_month:T", title="新客首次找師傅的月份"),
            y=alt.Y("回指率:Q", axis=alt.Axis(format="%"), title="回指率"),
            color=alt.Color("指標:N", title=""),
            tooltip=["cohort_month:T", "指標", alt.Tooltip("回指率:Q", format=".2%")],
        ).properties(height=280)
        st.altair_chart(line, use_container_width=True)

        # 近3期 vs 前3期
        def last_vs_prev(series):
            s = series.dropna()
            if len(s) < 6:
                return None
            last3 = s.iloc[-3:].mean()
            prev3 = s.iloc[-6:-3].mean()
            return last3 - prev3

        rate2_series = trend["rate2"].dropna()
        rate3_series = trend["rate3"].dropna()
        d2 = last_vs_prev(rate2_series)
        d3 = last_vs_prev(rate3_series)
        c1, c2 = st.columns(2)
        c1.metric("回指率(第2次) 近3期變化", f"{d2:+.2%}" if d2 is not None else "樣本不足")
        c2.metric("深度回指率(第3次) 近3期變化", f"{d3:+.2%}" if d3 is not None else "樣本不足")
    else:
        st.info("目前沒有足夠的回指資料繪製趨勢。")

st.subheader("師傅排行榜（Top 6，最近 3 個月）")
st.caption("僅計入滿 60 天的新客 cohort。")
col_good, col_watch = st.columns(2)

with col_good:
    st.markdown("<span style='color:#2ca02c;font-weight:700;'>表現較佳</span>", unsafe_allow_html=True)
    render_rank_table(
        score_df.dropna(subset=["score"]),
        "設計師",
        "score",
        "綜合分數最高",
        ascending=False,
        value_fmt="number1",
        top_n=6,
        text_color="#2ca02c",
    )
    render_rank_table(
        score_df.dropna(subset=["repeat_rate_3m"]),
        "設計師",
        "repeat_rate_3m",
        "回訪率最高(3M)",
        ascending=False,
        value_fmt="percent",
        top_n=6,
        text_color="#2ca02c",
    )
    if vacancy_by_designer is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_table(
            score_df.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最低(3M)",
            ascending=True,
            value_fmt="percent",
            top_n=6,
            text_color="#2ca02c",
        )
    render_rank_table(
        score_df.dropna(subset=["new_customers_3m"]),
        "設計師",
        "new_customers_3m",
        "新客數最多(3M)",
        ascending=False,
        value_fmt="int",
        top_n=6,
        text_color="#2ca02c",
    )

with col_watch:
    st.markdown("<span style='color:#d62728;font-weight:700;'>需關注</span>", unsafe_allow_html=True)
    render_rank_table(
        score_df.dropna(subset=["score"]),
        "設計師",
        "score",
        "綜合分數最低",
        ascending=True,
        value_fmt="number1",
        top_n=6,
        text_color="#d62728",
    )
    render_rank_table(
        score_df.dropna(subset=["repeat_rate_3m"]),
        "設計師",
        "repeat_rate_3m",
        "回訪率最低(3M)",
        ascending=True,
        value_fmt="percent",
        top_n=6,
        text_color="#d62728",
    )
    if vacancy_by_designer is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_table(
            score_df.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最高(3M)",
            ascending=False,
            value_fmt="percent",
            top_n=6,
            text_color="#d62728",
        )
    render_rank_table(
        score_df.dropna(subset=["new_customers_3m"]),
        "設計師",
        "new_customers_3m",
        "新客數最少(3M)",
        ascending=True,
        value_fmt="int",
        top_n=6,
        text_color="#d62728",
    )

st.subheader("回訪時間分布")
ret_series = compute_return_days(filtered_new_first, checkouts)
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

    hist_df = pd.DataFrame({"回訪天數": ret_series})
    hist = alt.Chart(hist_df).mark_bar(color="#4e79a7").encode(
        x=alt.X("回訪天數:Q", bin=alt.Bin(step=7), title="回訪天數(天)"),
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

# 師傅彙總（近3月回指）
st.markdown("**師傅彙總（近 3 月 / 回指）**")
if score_df is None or score_df.empty:
    st.info("目前沒有足夠的近 3 月回指資料。")
    designer_table = None
else:
    designer_table = score_df.rename(
        columns={
            "設計師": "師傅",
            "new_customers_3m": "新客數(3M)",
            "repeat_rate_3m": "回指率(第2次)",
            "deep_repeat_rate_3m": "深度回指率(第3次)",
            "vacancy_rate_3m": "空窗率(3M)",
            "score": "綜合分數",
        }
    )
    designer_table = designer_table[[c for c in designer_table.columns if c in [
        "師傅","新客數(3M)","回指率(第2次)","深度回指率(第3次)","空窗率(3M)","綜合分數"
    ]]]
    st.dataframe(designer_table.sort_values("綜合分數", ascending=False), use_container_width=True)

# 分店彙總
if has_store:
    st.markdown("**分店彙總**")
    store_table = summary_by_store.rename(
        columns={
            "分店": "分店",
            "matured_new_customers": "滿60天新客數",
            "churned": "流失數",
            "churn_rate": "流失率",
            "repeat_rate": "回訪率",
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
            "repeat_rate": "回訪率",
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
    "回訪率": overall["repeat_rate_matured"],
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
    file_name="new_customer_60d_churn_by_designer.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
