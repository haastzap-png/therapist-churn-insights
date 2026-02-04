import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
import re
from datetime import timedelta
from io import BytesIO

st.set_page_config(page_title="新客流失率分析", layout="wide")

st.title("新客流失率分析（按師傅）")

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
    chart_top_n = st.number_input("圖表顯示前 N 名（0=全部）", min_value=0, max_value=100, value=15, step=1)

st.write("""
本工具會：
- 預設以「帳單歷史首次結帳」判定新客（最穩）
- 使用帳單檔（預設：服務 + 票券）
- 以「首單後 60 天內未再結帳」計算流失
- Repeat Rate = 首單後 T 天內有回訪的新客比例
- 空窗率：依項目分鐘估算時長（1~30=0.5；31~60=1；61~90=1.5），月上限 168 小時
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

def render_bar_chart(df, category_col, value_col, title, color="#4e79a7", top_n=0):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return
    chart_df = df[[category_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=False)
    if top_n and len(chart_df) > top_n:
        chart_df = chart_df.head(int(top_n))
    chart_df[value_col] = chart_df[value_col].astype(float)
    height = min(520, 28 * len(chart_df) + 40)
    base = alt.Chart(chart_df).encode(
        y=alt.Y(f"{category_col}:N", sort="-x", title=""),
        x=alt.X(f"{value_col}:Q", axis=alt.Axis(format="%", title=title)),
        tooltip=[category_col, alt.Tooltip(value_col, format=".2%")],
    )
    bars = base.mark_bar(color=color)
    labels = base.mark_text(align="left", dx=4, color="#333").encode(
        text=alt.Text(f"{value_col}:Q", format=".1%")
    )
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)

def render_rank_table(df, name_col, value_col, title, ascending, value_fmt, top_n=3):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return
    rank_df = df[[name_col, value_col]].dropna().copy()
    rank_df = rank_df.sort_values(value_col, ascending=ascending).head(top_n)
    rank_df = rank_df.reset_index(drop=True)
    rank_df.insert(0, "名次", rank_df.index + 1)
    rank_df.rename(columns={name_col: "師傅", value_col: "數值"}, inplace=True)
    if value_fmt == "percent":
        rank_df["數值"] = rank_df["數值"].map(lambda v: f"{v:.2%}")
    elif value_fmt == "int":
        rank_df["數值"] = rank_df["數值"].map(lambda v: f"{int(v)}")
    st.markdown(f"**{title}**")
    st.table(rank_df)

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
        store_options = sorted([s for s in new_first["分店"].dropna().unique()])
        store_filter = st.multiselect("分店", store_options, default=store_options)
    else:
        store_filter = None
    designer_options = sorted([s for s in new_first["設計師"].dropna().unique()])
    designer_filter = st.multiselect("師傅", designer_options, default=designer_options)

filtered_new_first = new_first.copy()
if has_store and store_filter is not None:
    filtered_new_first = filtered_new_first[filtered_new_first["分店"].isin(store_filter)]
filtered_new_first = filtered_new_first[filtered_new_first["設計師"].isin(designer_filter)]

filtered_matured = filtered_new_first[filtered_new_first["matured"]].copy()

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

# Designer-level metrics for rankings
new_by_designer = (
    filtered_new_first.groupby("設計師")["phone_key"].count().reset_index(name="new_customers")
)
designer_metrics = summary_by_designer.merge(new_by_designer, on="設計師", how="outer")

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
    "new_customers": int(len(filtered_new_first)),
    "matured_new_customers": int(len(filtered_matured)),
    "churned_matured": int(filtered_matured["churn"].sum()),
}
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

st.subheader("摘要")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("新客數", overall["new_customers"])
col2.metric("成熟新客數", overall["matured_new_customers"])
col3.metric("流失人數", overall["churned_matured"])
col4.metric("流失率", f"{overall['churn_rate_matured']:.2%}" if overall["churn_rate_matured"] is not None else "-")
col5.metric("回訪率", f"{overall['repeat_rate_matured']:.2%}" if overall["repeat_rate_matured"] is not None else "-")

st.caption(f"資料截止時間：{end_date}")

st.subheader("各師傅流失率")
st.dataframe(summary_by_designer.sort_values("churn_rate", ascending=False), use_container_width=True)
st.subheader("各師傅流失率（圖）")
render_bar_chart(
    summary_by_designer.sort_values("churn_rate", ascending=False),
    "設計師",
    "churn_rate",
    "流失率",
    color="#e15759",
    top_n=chart_top_n,
)

st.subheader("各師傅回訪率（圖）")
render_bar_chart(
    summary_by_designer.sort_values("repeat_rate", ascending=False),
    "設計師",
    "repeat_rate",
    "回訪率",
    color="#59a14f",
    top_n=chart_top_n,
)

if has_store:
    st.subheader("各分店流失率")
    st.dataframe(summary_by_store.sort_values("churn_rate", ascending=False), use_container_width=True)
    st.subheader("各分店流失率（圖）")
    render_bar_chart(
        summary_by_store.sort_values("churn_rate", ascending=False),
        "分店",
        "churn_rate",
        "流失率",
        color="#e15759",
        top_n=chart_top_n,
    )
    st.subheader("各分店回訪率（圖）")
    render_bar_chart(
        summary_by_store.sort_values("repeat_rate", ascending=False),
        "分店",
        "repeat_rate",
        "回訪率",
        color="#59a14f",
        top_n=chart_top_n,
    )
    st.subheader("分店 x 師傅流失率")
    st.dataframe(summary_by_store_designer.sort_values("churn_rate", ascending=False), use_container_width=True)

st.subheader("師傅排行榜（Top 3）")
tab_pos, tab_neg = st.tabs(["正面榜", "負面榜"])

with tab_pos:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_rank_table(
            designer_metrics.dropna(subset=["churn_rate"]),
            "設計師",
            "churn_rate",
            "流失率最低",
            ascending=True,
            value_fmt="percent",
        )
    with col2:
        render_rank_table(
            designer_metrics.dropna(subset=["repeat_rate"]),
            "設計師",
            "repeat_rate",
            "回訪率最高",
            ascending=False,
            value_fmt="percent",
        )
    with col3:
        if vacancy_by_designer is None:
            st.info("無法計算空窗率（缺少項目分鐘）。")
        else:
            render_rank_table(
                designer_metrics.dropna(subset=["vacancy_rate"]),
                "設計師",
                "vacancy_rate",
                "空窗率最低",
                ascending=True,
                value_fmt="percent",
            )
    with col4:
        render_rank_table(
            designer_metrics.dropna(subset=["new_customers"]),
            "設計師",
            "new_customers",
            "新客數最多",
            ascending=False,
            value_fmt="int",
        )

with tab_neg:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        render_rank_table(
            designer_metrics.dropna(subset=["churn_rate"]),
            "設計師",
            "churn_rate",
            "流失率最高",
            ascending=False,
            value_fmt="percent",
        )
    with col2:
        render_rank_table(
            designer_metrics.dropna(subset=["repeat_rate"]),
            "設計師",
            "repeat_rate",
            "回訪率最低",
            ascending=True,
            value_fmt="percent",
        )
    with col3:
        if vacancy_by_designer is None:
            st.info("無法計算空窗率（缺少項目分鐘）。")
        else:
            render_rank_table(
                designer_metrics.dropna(subset=["vacancy_rate"]),
                "設計師",
                "vacancy_rate",
                "空窗率最高",
                ascending=False,
                value_fmt="percent",
            )
    with col4:
        render_rank_table(
            designer_metrics.dropna(subset=["new_customers"]),
            "設計師",
            "new_customers",
            "新客數最少",
            ascending=True,
            value_fmt="int",
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

st.subheader("流失名單")
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

display_cols = ["phone_key", name_col, "分店", "設計師", "結帳操作時間", "是否滿期", "是否流失"]
display_cols = [c for c in display_cols if c in detail_display.columns]
st.dataframe(detail_display[display_cols].sort_values("結帳操作時間"), use_container_width=True)

# Vacancy rate (monthly, 168h cap)
st.subheader("空窗率（月，168 小時上限）")
if vacancy_monthly is None:
    st.warning("帳單檔缺少 '項目' 欄位，無法估算服務時數與空窗率。")
else:
    display_vacancy = vacancy_monthly.copy()
    if has_store and store_filter is not None:
        display_vacancy = display_vacancy[display_vacancy["分店"].isin(store_filter)]
    display_vacancy = display_vacancy[display_vacancy["設計師"].isin(designer_filter)]
    st.dataframe(display_vacancy.sort_values(["month", "設計師"]), use_container_width=True)

# Download Excel
st.subheader("下載報表")

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    pd.DataFrame([overall]).to_excel(writer, index=False, sheet_name="summary")
    summary_by_designer.to_excel(writer, index=False, sheet_name="by_designer")
    if has_store:
        summary_by_store.to_excel(writer, index=False, sheet_name="by_store")
        summary_by_store_designer.to_excel(writer, index=False, sheet_name="by_store_designer")
    detail_display[display_cols].to_excel(writer, index=False, sheet_name="detail")
    if vacancy_monthly is not None:
        vacancy_monthly.to_excel(writer, index=False, sheet_name="vacancy_monthly")

st.download_button(
    label="下載 Excel",
    data=output.getvalue(),
    file_name="new_customer_60d_churn_by_designer.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
