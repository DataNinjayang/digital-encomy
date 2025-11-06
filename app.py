# -*- coding: utf-8 -*-
"""
中国上市公司数字化测评平台
本地桌面版 – 直接放“大表”文件夹
"""
# -------------------- 1. 标准库 --------------------
import warnings
import random
from pathlib import Path

# -------------------- 2. 第三方库 --------------------
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# -------------------- 3. 全局配置 --------------------
warnings.filterwarnings("ignore")
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# -------------------- 4. 路径：当前文件所在目录 --------------------
EXCEL_FILE = Path(__file__).with_name("中国上市企业数字化转型指数（2007-2020）(1).xlsx")

# -------------------- 5. 主题配色 --------------------
THEMES = {
    "砖红": {
        "font": "#B22222",
        "primary": "#B22222",
        "secondary": "#FF6F61",
        "bg": "#FFF5F5",
        "card": "#FFFFFF",
        "hover": "#FFE4E1",
        "text": "#2C2C2C",
        "accent": "#DC143C"
    }
}

# -------------------- 6. 会话状态 --------------------
if "theme" not in st.session_state:
    st.session_state.theme = "砖红"
if "page" not in st.session_state:
    st.session_state.page = "首页"

# -------------------- 7. 加载数据 --------------------
@st.cache_data(ttl=3600)
def load_data():
    if not EXCEL_FILE.exists():
        st.error(f"数据文件不存在，请确认路径：{EXCEL_FILE}")
        return pd.DataFrame()
    df = pd.read_excel(EXCEL_FILE)
    req = ["证券代码", "股票简称", "年份", "行业名称", "省份",
           "人工智能技术", "大数据技术", "云计算技术", "区块链技术", "数字化转型"]
    miss = [c for c in req if c not in df.columns]
    if miss:
        st.error(f"缺少列: {miss}")
        return pd.DataFrame()
    df = df.dropna(subset=["证券代码", "股票简称", "年份"])
    df["年份"] = pd.to_numeric(df["年份"], errors="coerce").astype("Int64")
    tech_cols = ["人工智能技术", "大数据技术", "云计算技术", "区块链技术"]
    for c in tech_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["技术总分"] = df[tech_cols].sum(axis=1)
    df["转型强度"] = pd.to_numeric(df["数字化转型"], errors="coerce").fillna(0)
    max_trans = df["转型强度"].max()
    if max_trans == 0:
        df["量化评分"] = 0.0
    else:
        df["量化评分"] = (df["转型强度"] / max_trans * 100).round(2)
    return df.sort_values(["证券代码", "年份"]).reset_index(drop=True)

# -------------------- 8. 可视化 --------------------
def trend_fig(df, code=None):
    if code:
        data = df[df["证券代码"] == code]
        name = data["股票简称"].iloc[0]
    else:
        data = df.groupby("年份")["量化评分"].mean().reset_index()
        name = "整体平均"
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=data["年份"], y=data["量化评分"],
        mode="lines+markers", name="评分",
        line=dict(color=THEMES[st.session_state.theme]["primary"], width=3)
    ))
    fig.update_layout(
        title=f"{name} 数字化转型趋势",
        xaxis_title="年份", yaxis_title="评分", height=400
    )
    return fig

def radar_fig(df, code):
    d = df[df["证券代码"] == code].iloc[-1]
    cats = ["人工智能技术", "大数据技术", "云计算技术", "区块链技术"]
    vals = [d[c] for c in cats]
    fig = go.Figure(go.Scatterpolar(
        r=vals, theta=cats, fill="toself",
        line_color=THEMES[st.session_state.theme]["primary"]
    ))
    fig.update_layout(
        title=f"{d['股票简称']} 技术维度",
        polar=dict(radialaxis=dict(range=[0, max(vals)*1.2])), height=400
    )
    return fig

# -------------------- 9. 页面 --------------------
def show_home(df):
    st.markdown(f"""
    <div style='text-align:center;padding:20px;background:linear-gradient(135deg,{THEMES[st.session_state.theme]["primary"]},{THEMES[st.session_state.theme]["secondary"]});color:white;border-radius:12px;'>
        <h1>🏢 中国上市公司数字化测评平台</h1>
        <p>2007-2020 年上市公司年报数据深度分析</p>
    </div>""", unsafe_allow_html=True)
    if df.empty:
        return
    total = df["证券代码"].nunique()
    st.markdown(f"""<div class='custom-card'>
        <h3>📈 数据概览</h3>
        <b>上市公司数量：</b>{total}<br>
        <b>年份跨度：</b>{df['年份'].min()} - {df['年份'].max()}<br>
        <b>记录条数：</b>{len(df)}
    </div>""", unsafe_allow_html=True)
    st.plotly_chart(trend_fig(df), use_container_width=True)

def show_company(df):
    st.markdown("### 🔍 企业分析")
    opts = [f"{c} - {df[df['证券代码']==c]['股票简称'].iloc[0]}" for c in sorted(df["证券代码"].unique())]
    # 随机企业按钮
    if st.button("🎲 随机企业"):
        st.session_state.rand_code = random.choice(list(df["证券代码"].unique()))
    if "rand_code" in st.session_state:
        default_idx = next(i for i, o in enumerate(opts) if str(st.session_state.rand_code) in o)
    else:
        default_idx = 0
    sel = st.selectbox("选择企业", opts, index=default_idx)
    code = int(sel.split(" - ")[0])
    comp = df[df["证券代码"] == code]
    if comp.empty:
        return
    st.markdown(f"#### {comp['股票简称'].iloc[0]}（{code}）")
    col1, col2 = st.columns(2)
    with col1:
        st.plotly_chart(trend_fig(df, code), use_container_width=True)
    with col2:
        st.plotly_chart(radar_fig(df, code), use_container_width=True)
    st.markdown("#### 历年数据")
    st.dataframe(comp[["年份", "量化评分", "技术总分"] + ["人工智能技术", "大数据技术", "云计算技术", "区块链技术"]].round(2))

def main():
    st.set_page_config(page_title="中国上市公司数字化测评平台", layout="wide")
    df = load_data()
    with st.sidebar:
        st.markdown("### 🎛️ 导航")
        page = st.radio("", ["首页", "企业分析"], index=0 if st.session_state.page == "首页" else 1)
        st.session_state.page = page
    if page == "首页":
        show_home(df)
    else:
        show_company(df)

if __name__ == "__main__":
    main()
