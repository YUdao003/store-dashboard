import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
from datetime import datetime, date, timedelta
import io

st.set_page_config(
    page_title="门店运营看板",
    page_icon="🏪",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 样式（浅色专业风）─────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f4f6fa; }
[data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e2e8f0; }
[data-testid="stSidebar"] * { color: #1a202c !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 { color: #1a202c !important; }

.block-container { padding-top: 1.5rem !important; }

.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.metric-label  { color: #64748b; font-size: 13px; margin-bottom: 6px; font-weight: 500; }
.metric-value  { color: #0f172a; font-size: 26px; font-weight: 700; line-height: 1.2; }
.metric-sub    { color: #94a3b8; font-size: 11px; margin-top: 4px; }

.section-title {
    color: #1e293b;
    font-size: 16px;
    font-weight: 700;
    margin: 24px 0 12px;
    padding-left: 10px;
    border-left: 4px solid #4f46e5;
}

div[data-testid="stTabs"] button {
    color: #64748b;
    font-size: 15px;
    font-weight: 500;
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: #4f46e5;
    border-bottom: 2px solid #4f46e5;
}

/* 主区域文字 */
h1, h2, h3, h4, p, label, span { color: #1e293b; }

/* 顶部信息条 */
.top-info {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 12px 20px;
    margin-bottom: 16px;
}
.top-info b { color: #4f46e5; }

/* 日报头部 */
.report-header {
    background: linear-gradient(135deg, #4f46e5, #6366f1);
    border-radius: 12px;
    padding: 20px 28px;
    margin-bottom: 16px;
    color: #ffffff;
}
.report-header h3 { color: #ffffff !important; margin: 0 0 8px; }
.report-header p, .report-header b { color: #ffffff !important; }
</style>
""", unsafe_allow_html=True)

PLATFORM_COLORS = {"美团": "#FFB300", "饿了么": "#1E88E5"}
CHART_BG = "#ffffff"
TEXT_COLOR = "#334155"
GRID_COLOR = "#e2e8f0"


# ── 数据加载 ──────────────────────────────────────────────────────────────────
def load_meituan(file) -> pd.DataFrame:
    if isinstance(file, str):
        df = pd.read_csv(file, encoding="gbk")
    else:
        raw = file.read()
        df = pd.read_csv(io.BytesIO(raw), encoding="gbk")
    df["日期"] = pd.to_datetime(df["日期"], format="%Y%m%d")
    df["平台"] = "美团"
    return df.rename(columns={
        "营业收入": "收入",
        "实付单均价": "单均实付",
        "入店人数": "进店人数",
        "入店转化率": "进店转化率",
        "入店新客": "新客进店人数",
        "入店老客": "老客进店人数",
        "入店次数": "进店次数",
        "平台服务费(含佣金和配送服务费)": "平台技术服务费",
        "顾客实付": "顾客实付总额",
    })


def load_eleme(file) -> pd.DataFrame:
    if isinstance(file, str):
        df = pd.read_excel(file, sheet_name="data", header=0)
    else:
        df = pd.read_excel(file, sheet_name="data", header=0)
    df["日期"] = pd.to_datetime(df["日期"])
    df["平台"] = "饿了么"
    return df


def load_all(mt_file, elm_file) -> pd.DataFrame:
    frames = []
    if mt_file is not None:
        try:
            frames.append(load_meituan(mt_file))
        except Exception as e:
            st.error(f"美团数据加载失败: {e}")
    if elm_file is not None:
        try:
            frames.append(load_eleme(elm_file))
        except Exception as e:
            st.error(f"饿了么数据加载失败: {e}")
    if not frames:
        return pd.DataFrame()

    shared_cols = [
        "日期", "门店名称", "平台", "收入", "有效订单",
        "单均实付", "顾客实付总额",
        "曝光人数", "进店人数", "下单转化率", "进店转化率",
        "曝光新客", "新客进店人数",
        "曝光老客", "老客进店人数",
        "曝光次数", "进店次数",
        "活动补贴", "平台技术服务费",
    ]
    dfs = []
    for d in frames:
        available = [c for c in shared_cols if c in d.columns]
        dfs.append(d[available].copy())
    combined = pd.concat(dfs, ignore_index=True)
    for col in ["收入", "有效订单", "单均实付", "曝光人数", "进店人数",
                "下单转化率", "进店转化率", "曝光新客", "新客进店人数",
                "曝光老客", "老客进店人数", "曝光次数", "进店次数"]:
        if col in combined.columns:
            combined[col] = pd.to_numeric(combined[col], errors="coerce").fillna(0)
    return combined


# ── 工具函数 ──────────────────────────────────────────────────────────────────
def fmt_money(v):
    if v >= 10000:
        return f"¥{v/10000:.2f}万"
    return f"¥{v:,.0f}"

def fmt_int(v):
    return f"{int(v):,}"

def fmt_pct(v):
    return f"{v*100:.1f}%"


def kpi_card(label, value, sub=""):
    sub_html = f'<div class="metric-sub">{sub}</div>' if sub else ""
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {sub_html}
    </div>
    """, unsafe_allow_html=True)


def _base_layout(fig, height=None):
    fig.update_layout(
        plot_bgcolor=CHART_BG, paper_bgcolor=CHART_BG,
        font=dict(color=TEXT_COLOR, family="-apple-system, Microsoft YaHei, sans-serif", size=12),
        legend=dict(orientation="h", y=1.08, x=0, bgcolor="rgba(0,0,0,0)"),
        margin=dict(l=10, r=20, t=40, b=10),
        hovermode="x unified",
    )
    if height:
        fig.update_layout(height=height)
    return fig


# ── 图表 ──────────────────────────────────────────────────────────────────────
def plot_daily_trend(df):
    daily = df.groupby(["日期", "平台"]).agg(
        收入=("收入", "sum"),
        订单=("有效订单", "sum"),
    ).reset_index()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    for plat in daily["平台"].unique():
        sub = daily[daily["平台"] == plat]
        color = PLATFORM_COLORS.get(plat, "#888")
        fig.add_trace(go.Scatter(
            x=sub["日期"], y=sub["收入"], name=f"{plat} 营业收入",
            line=dict(color=color, width=3),
            mode="lines+markers", marker=dict(size=8, line=dict(color="white", width=1.5)),
            hovertemplate=f"{plat} 收入: ¥%{{y:,.2f}}<extra></extra>",
        ), secondary_y=False)
        fig.add_trace(go.Bar(
            x=sub["日期"], y=sub["订单"], name=f"{plat} 订单量",
            marker_color=color, opacity=0.35,
            hovertemplate=f"{plat} 订单: %{{y:,.0f}}<extra></extra>",
        ), secondary_y=True)
    _base_layout(fig, height=420)
    fig.update_xaxes(gridcolor=GRID_COLOR, tickformat="%m-%d", showline=True, linecolor=GRID_COLOR)
    fig.update_yaxes(title_text="营业收入 (¥)", gridcolor=GRID_COLOR, tickformat=",.0f",
                     secondary_y=False, showline=True, linecolor=GRID_COLOR)
    fig.update_yaxes(title_text="订单量", gridcolor=GRID_COLOR, tickformat=",.0f",
                     secondary_y=True, showgrid=False)
    return fig


def _short_store_name(s):
    import re
    m = re.search(r"[（(](.+?)[)）]", s)
    return m.group(1) if m else s


def plot_store_rank(df, metric, label):
    store = df.groupby(["门店名称", "平台"]).agg(v=(metric, "sum")).reset_index()
    store = store.sort_values("v", ascending=True)
    store["门店短名"] = store["门店名称"].apply(_short_store_name)
    fig = px.bar(
        store, x="v", y="门店短名", color="平台",
        color_discrete_map=PLATFORM_COLORS,
        orientation="h", text="v",
        labels={"v": label, "门店短名": ""},
    )
    fig.update_traces(texttemplate="%{x:,.0f}", textposition="outside",
                      textfont=dict(color=TEXT_COLOR, size=11))
    _base_layout(fig, height=max(360, len(store) * 30))
    fig.update_xaxes(gridcolor=GRID_COLOR, tickformat=",.0f", showline=True, linecolor=GRID_COLOR)
    fig.update_yaxes(gridcolor=GRID_COLOR, automargin=True)
    fig.update_layout(bargap=0.35, margin=dict(l=10, r=80, t=40, b=10))
    return fig


def plot_funnel(df):
    exposure = df["曝光人数"].sum()
    visit    = df["进店人数"].sum()
    orders   = df["有效订单"].sum()
    fig = go.Figure(go.Funnel(
        y=["曝光人数", "进店人数", "下单人数"],
        x=[exposure, visit, orders],
        textinfo="value+percent initial",
        texttemplate="%{value:,.0f}<br>%{percentInitial}",
        textfont=dict(color="white", size=13),
        marker=dict(color=["#4f46e5", "#10b981", "#f59e0b"]),
        connector=dict(line=dict(color="#cbd5e1", width=2)),
    ))
    _base_layout(fig, height=380)
    fig.update_layout(margin=dict(l=10, r=10, t=20, b=10))
    return fig


def plot_new_old(df):
    new_exp = df["曝光新客"].sum() if "曝光新客" in df.columns else 0
    old_exp = df["曝光老客"].sum() if "曝光老客" in df.columns else 0
    new_vis = df["新客进店人数"].sum() if "新客进店人数" in df.columns else 0
    old_vis = df["老客进店人数"].sum() if "老客进店人数" in df.columns else 0

    fig = make_subplots(
        rows=1, cols=2,
        specs=[[{"type": "pie"}, {"type": "pie"}]],
        subplot_titles=["曝光：新客 vs 老客", "进店：新客 vs 老客"],
    )
    for i, (n, o) in enumerate([(new_exp, old_exp), (new_vis, old_vis)], 1):
        fig.add_trace(go.Pie(
            labels=["新客", "老客"], values=[n, o],
            marker=dict(colors=["#4f46e5", "#10b981"], line=dict(color="white", width=2)),
            hole=0.55,
            textinfo="label+percent",
            texttemplate="%{label}<br>%{value:,.0f}<br>(%{percent})",
            textfont=dict(color="white", size=12),
        ), row=1, col=i)
    _base_layout(fig, height=380)
    fig.update_layout(showlegend=False)
    for ann in fig["layout"]["annotations"]:
        ann["font"] = dict(color=TEXT_COLOR, size=13)
    return fig


def plot_conversion_by_store(df):
    conv = df.groupby(["门店名称", "平台"]).agg(
        进店转化=("进店转化率", "mean"),
        下单转化=("下单转化率", "mean"),
    ).reset_index()
    conv["门店短名"] = conv["门店名称"].apply(_short_store_name)
    conv = conv.sort_values("下单转化", ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name="进店转化率", y=conv["门店短名"], x=conv["进店转化"]*100,
        orientation="h", marker_color="#4f46e5", opacity=0.55,
        text=[f"{v*100:.1f}%" for v in conv["进店转化"]],
        textposition="outside", textfont=dict(color=TEXT_COLOR, size=10),
    ))
    fig.add_trace(go.Bar(
        name="下单转化率", y=conv["门店短名"], x=conv["下单转化"]*100,
        orientation="h", marker_color="#10b981",
        text=[f"{v*100:.1f}%" for v in conv["下单转化"]],
        textposition="outside", textfont=dict(color=TEXT_COLOR, size=10),
    ))
    _base_layout(fig, height=max(380, len(conv) * 32))
    fig.update_layout(barmode="group", margin=dict(l=10, r=80, t=40, b=10))
    fig.update_xaxes(gridcolor=GRID_COLOR, ticksuffix="%", showline=True, linecolor=GRID_COLOR)
    fig.update_yaxes(gridcolor=GRID_COLOR, automargin=True)
    return fig


# ── 预警计算 ──────────────────────────────────────────────────────────────────
ALERT_METRICS = {
    "营业收入":   {"col": "收入",        "fmt": lambda x: f"¥{x:,.2f}",        "low_is_bad": True},
    "有效订单":   {"col": "有效订单",     "fmt": lambda x: f"{int(x):,}",       "low_is_bad": True},
    "进店转化率": {"col": "进店转化率",   "fmt": lambda x: f"{x*100:.2f}%",     "low_is_bad": True},
    "下单转化率": {"col": "下单转化率",   "fmt": lambda x: f"{x*100:.2f}%",     "low_is_bad": True},
    "曝光人数":   {"col": "曝光人数",     "fmt": lambda x: f"{int(x):,}",       "low_is_bad": True},
    "进店人数":   {"col": "进店人数",     "fmt": lambda x: f"{int(x):,}",       "low_is_bad": True},
}


def compute_alerts(df, z_thresh=2.0, drop_thresh=0.30, monitor_metrics=None):
    """
    基于历史均值的异常检测：
      - Z-score：(今日值 - 历史均值) / 历史标准差，|z| > z_thresh 触发
      - 环比下跌：当指标"低=坏"时，下跌幅度 > drop_thresh 触发
    """
    if monitor_metrics is None:
        monitor_metrics = list(ALERT_METRICS.keys())

    alerts = []
    for (store, platform), grp in df.groupby(["门店名称", "平台"]):
        grp = grp.sort_values("日期").reset_index(drop=True)
        if len(grp) < 3:
            continue
        for i in range(len(grp)):
            today = grp.iloc[i]
            others = grp.drop(i)  # 除今日外所有历史数据
            if len(others) < 2:
                continue

            for name in monitor_metrics:
                info = ALERT_METRICS[name]
                col = info["col"]
                if col not in grp.columns:
                    continue
                today_val = today[col]
                hist_mean = others[col].mean()
                hist_std  = others[col].std()

                if hist_mean == 0 or pd.isna(hist_mean):
                    continue

                z = (today_val - hist_mean) / hist_std if hist_std and hist_std > 0 else 0
                deviation = (today_val - hist_mean) / hist_mean  # 相对偏离

                bad_direction = (deviation < 0) if info["low_is_bad"] else (deviation > 0)
                trigger_z    = abs(z) >= z_thresh and bad_direction
                trigger_drop = info["low_is_bad"] and (deviation <= -drop_thresh) and abs(deviation) > 0.05

                if trigger_z or trigger_drop:
                    if abs(z) >= 2.5 or abs(deviation) >= 0.5:
                        severity = "🔴 严重"
                        sev_rank = 3
                    elif abs(z) >= 2.0 or abs(deviation) >= 0.35:
                        severity = "🟠 警告"
                        sev_rank = 2
                    else:
                        severity = "🟡 注意"
                        sev_rank = 1

                    alerts.append({
                        "日期": today["日期"].date(),
                        "门店": store,
                        "平台": platform,
                        "指标": name,
                        "当日值": info["fmt"](today_val),
                        "历史均值": info["fmt"](hist_mean),
                        "偏离": f"{deviation*100:+.1f}%",
                        "Z值": f"{z:+.2f}",
                        "严重度": severity,
                        "_sev_rank": sev_rank,
                        "_date": today["日期"],
                        "_deviation": deviation,
                    })

    if not alerts:
        return pd.DataFrame()
    out = pd.DataFrame(alerts).sort_values(
        ["_sev_rank", "_date", "_deviation"],
        ascending=[False, False, True],
    )
    return out


def generate_daily_report(df, report_date) -> pd.DataFrame:
    day = df[df["日期"] == pd.Timestamp(report_date)]
    if day.empty:
        return pd.DataFrame()
    summary = day.groupby(["门店名称", "平台"]).agg(
        营业收入=("收入", "sum"),
        有效订单=("有效订单", "sum"),
        单均实付=("单均实付", "mean"),
        曝光人数=("曝光人数", "sum"),
        进店人数=("进店人数", "sum"),
        进店转化率=("进店转化率", "mean"),
        下单转化率=("下单转化率", "mean"),
    ).reset_index().sort_values("营业收入", ascending=False)
    summary["营业收入"] = summary["营业收入"].map(lambda x: f"¥{x:,.2f}")
    summary["单均实付"] = summary["单均实付"].map(lambda x: f"¥{x:.2f}")
    summary["进店转化率"] = summary["进店转化率"].map(lambda x: f"{x*100:.1f}%")
    summary["下单转化率"] = summary["下单转化率"].map(lambda x: f"{x*100:.1f}%")
    summary["曝光人数"] = summary["曝光人数"].map(lambda x: f"{int(x):,}")
    summary["进店人数"] = summary["进店人数"].map(lambda x: f"{int(x):,}")
    summary["有效订单"] = summary["有效订单"].map(lambda x: f"{int(x):,}")
    return summary.rename(columns={"门店名称": "门店"})


# ── 侧边栏 ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏪 门店运营看板")
    st.caption("美团 + 饿了么 数据监控")
    st.divider()

    st.markdown("#### 📁 数据文件")
    mt_upload  = st.file_uploader("美团 CSV", type=["csv"], key="mt")
    elm_upload = st.file_uploader("饿了么 Excel", type=["xlsx", "xls"], key="elm")

    # 本地调试时可选：从环境变量读取默认文件路径
    DEFAULT_MT  = os.environ.get("DASHBOARD_MT_FILE")
    DEFAULT_ELM = os.environ.get("DASHBOARD_ELM_FILE")

    mt_src  = mt_upload  if mt_upload  else (DEFAULT_MT  if DEFAULT_MT  and os.path.exists(DEFAULT_MT)  else None)
    elm_src = elm_upload if elm_upload else (DEFAULT_ELM if DEFAULT_ELM and os.path.exists(DEFAULT_ELM) else None)

    with st.spinner("加载数据..."):
        df_all = load_all(mt_src, elm_src)

    if df_all.empty:
        st.info("👆 请上传美团 CSV 和/或饿了么 Excel 文件开始使用")
        st.markdown("""
        <div style='background:#fff;border:1px solid #e2e8f0;border-radius:8px;
        padding:14px 18px;font-size:13px;color:#475569;margin-top:10px;'>
        <b style='color:#0f172a;'>📥 使用步骤：</b>
        <ol style='margin:6px 0 0 0;padding-left:20px;'>
        <li>从美团/饿了么商家后台导出门店数据</li>
        <li>点击上方上传按钮选择文件</li>
        <li>右侧自动生成看板和预警</li>
        </ol>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    st.success(f"✅ 已加载 {len(df_all):,} 条记录")

    st.divider()
    st.markdown("#### 🔍 筛选")

    all_dates = sorted(df_all["日期"].dt.date.unique())
    min_d, max_d = all_dates[0], all_dates[-1]

    # 健壮的日期范围处理（单选/双选都OK）
    raw_range = st.date_input(
        "日期范围",
        value=(min_d, max_d),
        min_value=min_d,
        max_value=max_d,
        key="date_range_input",
    )
    if isinstance(raw_range, (tuple, list)):
        if len(raw_range) == 2:
            start_d, end_d = raw_range[0], raw_range[1]
        elif len(raw_range) == 1:
            start_d = end_d = raw_range[0]
        else:
            start_d, end_d = min_d, max_d
    else:
        start_d = end_d = raw_range

    platforms = st.multiselect(
        "平台",
        df_all["平台"].unique().tolist(),
        default=df_all["平台"].unique().tolist(),
    )

    all_stores = sorted(df_all["门店名称"].unique().tolist())
    selected_stores = st.multiselect("门店（留空=全部）", all_stores)

    st.divider()
    st.markdown("#### ⚠️ 预警阈值")
    z_thresh = st.slider(
        "Z-score 阈值",
        min_value=1.0, max_value=3.0, value=2.0, step=0.1,
        help="今日值偏离历史均值多少个标准差时触发。建议 1.5~2.5。",
    )
    drop_thresh = st.slider(
        "环比下跌阈值",
        min_value=0.1, max_value=0.6, value=0.30, step=0.05,
        format="%.0f%%",
        help="对比历史均值，下跌超过该比例触发预警。",
    )
    monitor_metrics = st.multiselect(
        "监控指标",
        list(ALERT_METRICS.keys()),
        default=["营业收入", "进店转化率", "下单转化率", "曝光人数"],
    )


# ── 过滤数据 ──────────────────────────────────────────────────────────────────
if not platforms:
    st.warning("⚠️ 请至少选择一个平台")
    st.stop()

mask = (
    (df_all["日期"].dt.date >= start_d) &
    (df_all["日期"].dt.date <= end_d) &
    (df_all["平台"].isin(platforms))
)
if selected_stores:
    mask &= df_all["门店名称"].isin(selected_stores)
df = df_all[mask].copy()

if df.empty:
    st.warning("⚠️ 当前筛选条件下无数据，请调整筛选条件")
    st.stop()


# ── 顶部信息条 ────────────────────────────────────────────────────────────────
st.markdown(f"""
<h1 style='font-size:24px;margin-bottom:8px;color:#0f172a;'>📊 门店运营监控看板</h1>
<div class="top-info" style='font-size:13px;color:#475569;'>
    📅 周期 <b>{start_d} ~ {end_d}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
    🏷 平台 <b>{"、".join(platforms)}</b> &nbsp;&nbsp;|&nbsp;&nbsp;
    🏪 门店数 <b>{df["门店名称"].nunique()} 家</b> &nbsp;&nbsp;|&nbsp;&nbsp;
    📦 记录数 <b>{len(df):,}</b>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab_alert, tab4 = st.tabs(
    ["📈 综合概览", "🏪 门店分析", "👥 流量分析", "⚠️ 预警监控", "📋 日报"]
)

# ── Tab 1: 综合概览 ───────────────────────────────────────────────────────────
with tab1:
    total_rev  = df["收入"].sum()
    total_ord  = df["有效订单"].sum()
    avg_price  = (total_rev / total_ord) if total_ord else 0
    total_expo = df["曝光人数"].sum() if "曝光人数" in df.columns else 0
    total_vis  = df["进店人数"].sum() if "进店人数" in df.columns else 0
    overall_conv = (total_ord / total_vis) if total_vis else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi_card("总营业收入", fmt_money(total_rev), f"原值 ¥{total_rev:,.2f}")
    with c2: kpi_card("总有效订单", fmt_int(total_ord), "单")
    with c3: kpi_card("加权客单价", f"¥{avg_price:.2f}", "= 总收入 / 总订单")
    with c4: kpi_card("总曝光人数", fmt_int(total_expo), "人")
    with c5: kpi_card("整体下单转化", fmt_pct(overall_conv), "= 订单 / 进店")

    st.markdown('<div class="section-title">每日营业收入 & 订单量趋势</div>', unsafe_allow_html=True)
    st.plotly_chart(plot_daily_trend(df), width='stretch')

    st.markdown('<div class="section-title">各平台汇总</div>', unsafe_allow_html=True)
    plat_summary = df.groupby("平台").agg(
        营业收入=("收入", "sum"),
        有效订单=("有效订单", "sum"),
        曝光人数=("曝光人数", "sum"),
        进店人数=("进店人数", "sum"),
    ).reset_index()
    plat_summary["客单价"] = plat_summary["营业收入"] / plat_summary["有效订单"].replace(0, 1)
    plat_summary["下单转化率"] = plat_summary["有效订单"] / plat_summary["进店人数"].replace(0, 1)
    plat_summary["营业收入"]  = plat_summary["营业收入"].map(lambda x: f"¥{x:,.2f}")
    plat_summary["有效订单"]  = plat_summary["有效订单"].map(lambda x: f"{int(x):,}")
    plat_summary["曝光人数"]  = plat_summary["曝光人数"].map(lambda x: f"{int(x):,}")
    plat_summary["进店人数"]  = plat_summary["进店人数"].map(lambda x: f"{int(x):,}")
    plat_summary["客单价"]    = plat_summary["客单价"].map(lambda x: f"¥{x:.2f}")
    plat_summary["下单转化率"]= plat_summary["下单转化率"].map(lambda x: f"{x*100:.1f}%")
    st.dataframe(plat_summary, width='stretch', hide_index=True)


# ── Tab 2: 门店分析 ───────────────────────────────────────────────────────────
with tab2:
    col_left, col_right = st.columns(2)
    with col_left:
        st.markdown('<div class="section-title">门店营业收入排名</div>', unsafe_allow_html=True)
        st.plotly_chart(plot_store_rank(df, "收入", "营业收入 (¥)"), width='stretch')
    with col_right:
        st.markdown('<div class="section-title">门店订单量排名</div>', unsafe_allow_html=True)
        st.plotly_chart(plot_store_rank(df, "有效订单", "有效订单 (单)"), width='stretch')

    st.markdown('<div class="section-title">门店明细数据</div>', unsafe_allow_html=True)
    store_detail = df.groupby(["门店名称", "平台"]).agg(
        营业收入=("收入", "sum"),
        有效订单=("有效订单", "sum"),
        曝光人数=("曝光人数", "sum"),
        进店人数=("进店人数", "sum"),
    ).reset_index().sort_values("营业收入", ascending=False)
    store_detail["客单价"] = store_detail["营业收入"] / store_detail["有效订单"].replace(0, 1)
    store_detail["下单转化率"] = store_detail["有效订单"] / store_detail["进店人数"].replace(0, 1)
    store_detail["营业收入"]   = store_detail["营业收入"].map(lambda x: f"¥{x:,.2f}")
    store_detail["有效订单"]   = store_detail["有效订单"].map(lambda x: f"{int(x):,}")
    store_detail["曝光人数"]   = store_detail["曝光人数"].map(lambda x: f"{int(x):,}")
    store_detail["进店人数"]   = store_detail["进店人数"].map(lambda x: f"{int(x):,}")
    store_detail["客单价"]     = store_detail["客单价"].map(lambda x: f"¥{x:.2f}")
    store_detail["下单转化率"] = store_detail["下单转化率"].map(lambda x: f"{x*100:.1f}%")
    st.dataframe(store_detail, width='stretch', hide_index=True)


# ── Tab 3: 流量分析（按平台拆分）──────────────────────────────────────────────
with tab3:
    platforms_in_df = df["平台"].unique().tolist()

    if len(platforms_in_df) == 0:
        st.info("当前筛选下无数据")
    else:
        # 每个平台一列
        cols = st.columns(len(platforms_in_df))
        for i, plat in enumerate(platforms_in_df):
            sub = df[df["平台"] == plat]
            with cols[i]:
                color = PLATFORM_COLORS.get(plat, "#888")
                st.markdown(f"""
                <div style='background:{color}1A;border-left:4px solid {color};
                padding:10px 16px;border-radius:6px;margin-bottom:8px;'>
                    <b style='color:#0f172a;font-size:16px;'>🏷 {plat}</b>
                    <span style='color:#64748b;font-size:12px;margin-left:10px;'>
                    {sub['门店名称'].nunique()} 家门店 · {len(sub):,} 条记录
                    </span>
                </div>
                """, unsafe_allow_html=True)

                # 平台级 KPI
                p_exp = sub["曝光人数"].sum()
                p_vis = sub["进店人数"].sum()
                p_ord = sub["有效订单"].sum()
                p_rev = sub["收入"].sum()
                k1, k2, k3 = st.columns(3)
                with k1: kpi_card("总营业额", fmt_money(p_rev))
                with k2: kpi_card("曝光→进店", f"{(p_vis/p_exp*100) if p_exp else 0:.2f}%")
                with k3: kpi_card("进店→下单", f"{(p_ord/p_vis*100) if p_vis else 0:.2f}%")

                st.markdown(f'<div class="section-title">{plat} · 漏斗</div>', unsafe_allow_html=True)
                st.plotly_chart(plot_funnel(sub), width='stretch', key=f"funnel_{plat}")

                st.markdown(f'<div class="section-title">{plat} · 新客 vs 老客</div>', unsafe_allow_html=True)
                st.plotly_chart(plot_new_old(sub), width='stretch', key=f"newold_{plat}")

                st.markdown(f'<div class="section-title">{plat} · 各门店转化率</div>', unsafe_allow_html=True)
                st.plotly_chart(plot_conversion_by_store(sub), width='stretch', key=f"conv_{plat}")


# ── Tab 预警: 预警监控 ────────────────────────────────────────────────────────
with tab_alert:
    if not monitor_metrics:
        st.info("请在左侧选择至少一个监控指标")
    else:
        alerts_df = compute_alerts(df, z_thresh=z_thresh, drop_thresh=drop_thresh,
                                   monitor_metrics=monitor_metrics)

        # 概览统计
        if alerts_df.empty:
            sev_count = {"🔴 严重": 0, "🟠 警告": 0, "🟡 注意": 0}
        else:
            sev_count = alerts_df["严重度"].value_counts().to_dict()
        c1, c2, c3, c4 = st.columns(4)
        with c1: kpi_card("🔴 严重", str(sev_count.get("🔴 严重", 0)), "高度异常")
        with c2: kpi_card("🟠 警告", str(sev_count.get("🟠 警告", 0)), "明显偏离")
        with c3: kpi_card("🟡 注意", str(sev_count.get("🟡 注意", 0)), "轻度异常")
        with c4: kpi_card("✅ 监控指标", str(len(monitor_metrics)), f"Z阈值={z_thresh}, 下跌阈值={int(drop_thresh*100)}%")

        st.markdown('<div class="section-title">预警逻辑说明</div>', unsafe_allow_html=True)
        st.markdown(f"""
<div style='background:#fff;border:1px solid #e2e8f0;border-radius:8px;padding:14px 18px;font-size:13px;color:#475569;'>
对每家门店在每个平台上的每个指标：
<ul style='margin:8px 0 0 0;padding-left:20px;'>
<li><b>Z-score 检测</b>：当日值偏离该门店历史均值 ≥ <b>{z_thresh}σ</b> 且方向不利时触发</li>
<li><b>环比下跌检测</b>：当日值低于历史均值 ≥ <b>{int(drop_thresh*100)}%</b> 时触发</li>
<li><b>严重度分级</b>：|Z|≥2.5 或下跌≥50% → 🔴 严重　|　|Z|≥2.0 或下跌≥35% → 🟠 警告　|　其它 → 🟡 注意</li>
</ul>
</div>
""", unsafe_allow_html=True)

        if alerts_df.empty:
            st.success("✅ 当前筛选范围内未发现明显异常")
        else:
            st.markdown('<div class="section-title">预警明细</div>', unsafe_allow_html=True)
            # 筛选器
            f1, f2 = st.columns([1, 2])
            with f1:
                sev_filter = st.multiselect(
                    "严重度",
                    ["🔴 严重", "🟠 警告", "🟡 注意"],
                    default=["🔴 严重", "🟠 警告", "🟡 注意"],
                    key="sev_filter",
                )
            with f2:
                metric_filter = st.multiselect(
                    "指标",
                    alerts_df["指标"].unique().tolist(),
                    default=alerts_df["指标"].unique().tolist(),
                    key="metric_filter",
                )

            show_df = alerts_df[
                alerts_df["严重度"].isin(sev_filter) &
                alerts_df["指标"].isin(metric_filter)
            ][["日期", "门店", "平台", "指标", "当日值", "历史均值", "偏离", "Z值", "严重度"]]

            if show_df.empty:
                st.info("当前筛选下无预警")
            else:
                st.dataframe(show_df, width='stretch', hide_index=True, height=420)

                # 导出
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    show_df.to_excel(writer, index=False, sheet_name="预警明细")
                buf.seek(0)
                st.download_button(
                    "⬇️ 下载预警 Excel",
                    data=buf.getvalue(),
                    file_name=f"预警明细_{start_d}_{end_d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_alerts",
                )


# ── Tab 4: 日报 ───────────────────────────────────────────────────────────────
with tab4:
    st.markdown('<div class="section-title">选择日期生成日报</div>', unsafe_allow_html=True)

    report_date = st.date_input(
        "报告日期",
        value=max_d,
        min_value=min_d,
        max_value=max_d,
        key="report_date_input",
    )

    if isinstance(report_date, (tuple, list)):
        report_date = report_date[0]

    report_df = generate_daily_report(df_all, report_date)

    if report_df.empty:
        st.warning(f"⚠️ {report_date} 无数据")
    else:
        day_data = df_all[df_all["日期"] == pd.Timestamp(report_date)]
        total_rev_day = day_data["收入"].sum()
        total_ord_day = day_data["有效订单"].sum()
        total_exp_day = day_data["曝光人数"].sum() if "曝光人数" in day_data.columns else 0
        store_cnt_day = day_data["门店名称"].nunique()
        avg_price_day = (total_rev_day / total_ord_day) if total_ord_day else 0

        st.markdown(f"""
<div class="report-header">
<h3>📋 {report_date} 经营日报</h3>
<p style='margin:0;font-size:14px;'>
  营业收入 <b>{fmt_money(total_rev_day)}</b> &nbsp;|&nbsp;
  有效订单 <b>{fmt_int(total_ord_day)} 单</b> &nbsp;|&nbsp;
  客单价 <b>¥{avg_price_day:.2f}</b> &nbsp;|&nbsp;
  曝光人数 <b>{fmt_int(total_exp_day)}</b> &nbsp;|&nbsp;
  活跃门店 <b>{store_cnt_day} 家</b>
</p>
</div>
""", unsafe_allow_html=True)

        st.dataframe(report_df, width='stretch', hide_index=True)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            report_df.to_excel(writer, index=False, sheet_name=str(report_date))
        buf.seek(0)
        st.download_button(
            label="⬇️ 下载日报 Excel",
            data=buf.getvalue(),
            file_name=f"日报_{report_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{report_date}",
        )
