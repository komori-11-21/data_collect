import streamlit as st
import pandas as pd
import os
import glob
import io
import plotly.express as px

st.set_page_config(
    page_title="全景功耗监控大屏",
    page_icon="🌌",
    layout="wide"
)

st.markdown("""
    <style>
        /* 强制隐藏 Form 内文本输入框自带的 "Press Enter to submit form" 英文提示 */
        div[data-testid="InputInstructions"] {
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# 🛠️ 核心硬编码配置区 (Hardcoded Configuration)
# =====================================================================

DATA_DIR = "./data_source"
JOIN_KEYS = ["id", "data"]

# 1. 静态数值类表（挂载到上半屏的表格）
TABLE_CONFIGS = [
    {
        "keyword": "standby",
        "type": "wide",
        "label": "待机",
        "val_col": "待机功耗",
        "time_col": "待机时长",
        "val_unit": "mA",
        "time_unit": "s"
    },
    {
        "keyword": "apps",
        "type": "long",
        "item_col": "应用名称",
        "val_col": "功耗值",
        "time_col": "使用时长",
        "val_unit": "mA",
        "time_unit": "s"
    },
    {
        "keyword": "volume",
        "type": "long",
        "item_col": "媒体种类",
        "val_col": "音量值",
        "time_col": "持续时长",
        "val_unit": "%",
        "time_unit": "s",
        "dedup_rule": "longest_duration"
    },
    {
        "keyword": "sleep",
        "type": "wide",
        "label": "休眠",
        "val_col": "休眠次数",
        "val_unit": "次"
    },
    {
        "keyword": "events",
        "type": "event",
        "item_col": "状态名",
        "action_col": "动作"
    },
    {
        "keyword": "app_detail",
        "type": "app_detail",
        "app_col": "应用名称",
        "duration_col": "运行时长",
        "power_col": "功耗值"
    }
]

# =====================================================================

# --- 样式设计 ---
st.markdown("""
    <style>
    .main { background-color: #f4f6f9; }
    .title-box {
        padding: 30px;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 12px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.1);
        margin-bottom: 25px;
        text-align: center;
        color: white;
    }
    .title-box h2 {
        color: white !important;
        margin: 0;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 600;
        letter-spacing: 1px;
    }
    /* 强制隐藏 Form 内文本输入框自带的英文提示 */
    div[data-testid="InputInstructions"] {
        display: none !important;
    }
    /* 调整 Metric 卡片样式 */
    div[data-testid="metric-container"] {
        background-color: white;
        padding: 15px 20px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        border-left: 4px solid #2a5298;
    }
    /* ==================================
       侧边栏纯净导航 CSS 深度优化
       ================================== */
    [data-testid="stSidebar"] [role="radiogroup"] label {
        padding: 12px 15px;
        background-color: transparent;
        border-radius: 8px;
        margin-bottom: 5px;
        transition: all 0.2s ease;
        cursor: pointer;
    }
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {
        background-color: #f0f2f6;
        transform: translateX(4px);
    }
    /* 隐藏原生的圆形 radio button 节点 */
    [data-testid="stSidebar"] [role="radiogroup"] label > div:first-child {
        display: none !important;
    }
    /* 调整文字 */
    [data-testid="stSidebar"] [role="radiogroup"] div[data-testid="stMarkdownContainer"] p {
        font-size: 16px;
        font-weight: 500;
        margin: 0;
    }
    /* 选中时的极光高亮样式 */
    [data-testid="stSidebar"] [role="radiogroup"] label[aria-checked="true"] {
        background-color: #e6f0fa !important;
        border-left: 4px solid #1e3c72 !important;
        color: #1e3c72 !important;
    }
    [data-testid="stSidebar"] [role="radiogroup"] label[aria-checked="true"] p {
        color: #1e3c72 !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- 模拟数据生成 ---
def ensure_mock_data():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    dates = ["2026-05-01", "2026-05-02"]
    ids = ["DEV_001", "DEV_002"]

    # 丰富待机和休眠数据，补充更多交叉项
    standby_data = [
        {"id": "DEV_001", "data": "2026-05-01", "待机功耗": 15, "待机时长": 120},
        {"id": "DEV_001", "data": "2026-05-02", "待机功耗": 18, "待机时长": 140},
        {"id": "DEV_002", "data": "2026-05-01", "待机功耗": 12, "待机时长": 110},
        {"id": "DEV_002", "data": "2026-05-02", "待机功耗": 10, "待机时长": 150}
    ]
    pd.DataFrame(standby_data).to_excel(os.path.join(DATA_DIR, "table_1_standby.xlsx"), index=False)
    
    sleep_data = [
        {"id": "DEV_001", "data": "2026-05-01", "休眠次数": 5},
        {"id": "DEV_001", "data": "2026-05-02", "休眠次数": 7},
        {"id": "DEV_002", "data": "2026-05-01", "休眠次数": 3},
        {"id": "DEV_002", "data": "2026-05-02", "休眠次数": 8}
    ]
    pd.DataFrame(sleep_data).to_excel(os.path.join(DATA_DIR, "table_2_sleep.xlsx"), index=False)

    # 丰富音量数据 (增加多种类型的音量状态切换测试 longest_duration)
    volume_data = [
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 08:00:00", "媒体种类": "铃声音量", "音量值": 100, "持续时长": 1000},
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 09:00:00", "媒体种类": "铃声音量", "音量值": 50, "持续时长": 52000},
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 23:30:00", "媒体种类": "铃声音量", "音量值": 20, "持续时长": 1800},
        
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 10:00:00", "媒体种类": "媒体音量", "音量值": 80, "持续时长": 7200},
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 12:00:00", "媒体种类": "媒体音量", "音量值": 0, "持续时长": 43200},
        
        {"id": "DEV_002", "data": "2026-05-01", "发生时间": "2026-05-01 07:00:00", "媒体种类": "闹钟音量", "音量值": 100, "持续时长": 300},
        {"id": "DEV_002", "data": "2026-05-01", "发生时间": "2026-05-01 07:05:00", "媒体种类": "闹钟音量", "音量值": 0, "持续时长": 86100}
    ]
    pd.DataFrame(volume_data).to_excel(os.path.join(DATA_DIR, "table_3_volume.xlsx"), index=False)

    # 丰富应用数据，多造一点不同类型和时长的 APP
    app_data = [
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "微信", "功耗值": 50, "使用时长": 180},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "抖音", "功耗值": 80, "使用时长": 240},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "王者荣耀", "功耗值": 150, "使用时长": 3600},
        
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "微信", "功耗值": 45, "使用时长": 150},
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "Bilibili", "功耗值": 120, "使用时长": 1800},
        
        {"id": "DEV_002", "data": "2026-05-01", "应用名称": "淘宝", "功耗值": 60, "使用时长": 300},
        {"id": "DEV_002", "data": "2026-05-02", "应用名称": "高德地图", "功耗值": 200, "使用时长": 7200}
    ]
    pd.DataFrame(app_data).to_excel(os.path.join(DATA_DIR, "table_4_apps.xlsx"), index=False)

    # 事件表：模拟多天高频打印的状态轨迹
    event_data = [
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 08:00:00", "状态名": "省电模式", "动作": "1开机的时候开"},
        # 模拟心跳冗余打印 (应当被去抖算法剔除)
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 09:00:00", "状态名": "省电模式", "动作": "1开机的时候开"},
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 10:00:00", "状态名": "省电模式", "动作": "1开机的时候开"},
        
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 12:30:00", "状态名": "省电模式", "动作": "4用户手动关"},
        {"id": "DEV_001", "data": "2026-05-01", "发生时间": "2026-05-01 18:00:00", "状态名": "省电模式", "动作": "3用户手动开"},
        
        {"id": "DEV_001", "data": "2026-05-02", "发生时间": "2026-05-02 08:00:00", "状态名": "省电模式", "动作": "1开机的时候开"},
        
        {"id": "DEV_002", "data": "2026-05-01", "发生时间": "2026-05-01 07:00:00", "状态名": "蓝牙状态", "动作": "3用户手动开"},
        {"id": "DEV_002", "data": "2026-05-01", "发生时间": "2026-05-01 07:05:00", "状态名": "蓝牙状态", "动作": "4用户手动关"}
    ]
    pd.DataFrame(event_data).to_excel(os.path.join(DATA_DIR, "table_5_events.xlsx"), index=False)

    # 应用明细表：每个应用每天多次运行的时长和功耗
    app_detail_data = [
        # DEV_001 2026-05-01
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "微信", "运行时长": 35, "功耗值": 8},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "微信", "运行时长": 120, "功耗值": 25},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "微信", "运行时长": 5, "功耗值": 2},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "抹音", "运行时长": 200, "功耗值": 65},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "抹音", "运行时长": 15, "功耗值": 5},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "王者荣耀", "运行时长": 1800, "功耗值": 120},
        {"id": "DEV_001", "data": "2026-05-01", "应用名称": "王者荣耀", "运行时长": 2400, "功耗值": 150},
        # DEV_001 2026-05-02
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "微信", "运行时长": 90, "功耗值": 18},
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "微信", "运行时长": 60, "功耗值": 12},
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "Bilibili", "运行时长": 900, "功耗值": 80},
        {"id": "DEV_001", "data": "2026-05-02", "应用名称": "Bilibili", "运行时长": 300, "功耗值": 30},
        # DEV_002 2026-05-01
        {"id": "DEV_002", "data": "2026-05-01", "应用名称": "淘宝", "运行时长": 180, "功耗值": 40},
        {"id": "DEV_002", "data": "2026-05-01", "应用名称": "淘宝", "运行时长": 60, "功耗值": 15},
        {"id": "DEV_002", "data": "2026-05-01", "应用名称": "淘宝", "运行时长": 8, "功耗值": 3},
        # DEV_002 2026-05-02
        {"id": "DEV_002", "data": "2026-05-02", "应用名称": "高德地图", "运行时长": 3600, "功耗值": 180},
        {"id": "DEV_002", "data": "2026-05-02", "应用名称": "高德地图", "运行时长": 1200, "功耗值": 55},
    ]
    pd.DataFrame(app_detail_data).to_excel(os.path.join(DATA_DIR, "table_6_app_detail.xlsx"), index=False)


ensure_mock_data()


# --- 核心数据加载与分离逻辑 ---
# @st.cache_data(show_spinner=False) # 暂时注释缓存，防止读不到新文件
def load_data():
    all_files = glob.glob(os.path.join(DATA_DIR, "*.xls*"))
    standard_dfs = []
    event_df_list = []

    for file_path in all_files:
        file_name = os.path.basename(file_path).lower()
        try:
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
            if 'data' in df.columns:
                df['data'] = pd.to_datetime(df['data']).dt.strftime('%Y-%m-%d')
        except Exception:
            continue

        # 核心解析拼装
        matched_config = next((cfg for cfg in TABLE_CONFIGS if cfg["keyword"].lower() in file_name), None)
        if not matched_config: continue

        p_keys = [k for k in JOIN_KEYS if k in df.columns]
        order = TABLE_CONFIGS.index(matched_config)

        if matched_config["type"] == "wide":
            sub_df = df.copy()
            label = matched_config["label"]
            v_col = matched_config.get("val_col")
            t_col = matched_config.get("time_col")
            v_unit = matched_config.get("val_unit", "")
            t_unit = matched_config.get("time_unit", "")
            
            unit_suffix = ""
            if v_unit and t_unit:
                unit_suffix = f" ({v_unit} | {t_unit})"
            elif v_unit:
                unit_suffix = f" ({v_unit})"
            
            sub_df['指标分类'] = label + unit_suffix
            sub_df['数值'] = sub_df[v_col] if v_col in sub_df.columns else ""
            sub_df['时长'] = sub_df[t_col] if t_col and t_col in sub_df.columns else ""
            
            sub_df = sub_df[p_keys + ['指标分类', '时长', '数值']]
            sub_df['排序层级'] = order
            standard_dfs.append(sub_df)

        elif matched_config["type"] == "long":
            item_col = matched_config.get("item_col")
            v_col = matched_config.get("val_col")
            t_col = matched_config.get("time_col")
            v_unit = matched_config.get("val_unit", "")
            t_unit = matched_config.get("time_unit", "")
            
            unit_suffix = ""
            if v_unit and t_unit:
                unit_suffix = f" ({v_unit} | {t_unit})"
            elif v_unit:
                unit_suffix = f" ({v_unit})"

            if item_col in df.columns:
                sub_df = df.copy()

                # 【新增核心逻辑】：业务停留时长最长的去重机制
                if matched_config.get("dedup_rule") == "longest_duration":
                    if t_col and t_col in sub_df.columns:
                        # 优先：如果有真实的持续时长列，直接基于它取最长
                        sub_df['duration_tmp'] = pd.to_numeric(sub_df[t_col], errors='coerce').fillna(0)
                        idx_max = sub_df.groupby(p_keys + [item_col])['duration_tmp'].idxmax()
                        sub_df = sub_df.loc[idx_max].copy()
                    elif '发生时间' in sub_df.columns:
                        # 兜底：如果只有发生时间流，用时间差推算
                        sub_df['发生时间'] = pd.to_datetime(sub_df['发生时间'])
                        sub_df = sub_df.sort_values(by=p_keys + [item_col, '发生时间'])

                        # 计算每次状态的维持时间 (当前条与下一条的时间差)
                        sub_df['next_time'] = sub_df.groupby(p_keys + [item_col])['发生时间'].shift(-1)
                        # 最后一条的 next_time 默认补全为当天深夜 23:59:59
                        data_end = pd.to_datetime(sub_df['data']) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                        sub_df['next_time'] = sub_df['next_time'].fillna(data_end)
                        sub_df['duration_tmp'] = (sub_df['next_time'] - sub_df['发生时间']).dt.total_seconds()

                        # 只保留 duration 最大的那一行！
                        idx_max = sub_df.groupby(p_keys + [item_col])['duration_tmp'].idxmax()
                        sub_df = sub_df.loc[idx_max].copy()

                sub_df['指标分类'] = " └─ " + sub_df[item_col].astype(str) + unit_suffix
                sub_df['数值'] = sub_df[v_col] if v_col and v_col in sub_df.columns else ""
                sub_df['时长'] = sub_df[t_col] if t_col and t_col in sub_df.columns else ""

                sub_df = sub_df[p_keys + ['指标分类', '时长', '数值']]
                sub_df['排序层级'] = order
                standard_dfs.append(sub_df)

        elif matched_config["type"] == "event":
            item_col = matched_config.get("item_col")
            act_col = matched_config.get("action_col")
            
            if item_col in df.columns and act_col in df.columns:
                sub_df = df.copy()
                
                # 【神级数据压缩】：将高频离散事件流聚合为按天的状态轨迹
                if '发生时间' in sub_df.columns:
                    # 按照时间严格排序
                    sub_df['发生时间'] = pd.to_datetime(sub_df['发生时间'])
                    sub_df = sub_df.sort_values(by=p_keys + [item_col, '发生时间'])
                    
                    # === 状态流去抖 (Debounce)：剔除连续重复的心跳日志 ===
                    sub_df['prev_act'] = sub_df.groupby(p_keys + [item_col])[act_col].shift(1)
                    sub_df = sub_df[sub_df[act_col] != sub_df['prev_act']].copy()
                    # ====================================================

                    # 提取时分秒，并和动作拼接成轨迹点
                    time_str = sub_df['发生时间'].dt.strftime('%H:%M:%S')
                    sub_df['轨迹点'] = "[" + time_str + "] " + sub_df[act_col].astype(str)
                    
                    # 按天聚合，用 ➜ 符号连接，压扁为单格长文本
                    agg_df = sub_df.groupby(p_keys + [item_col])['轨迹点'].apply(lambda x: " ➜ ".join(x)).reset_index()
                    agg_df['数值'] = agg_df['轨迹点']
                else:
                    # 兜底：没有时间戳则简单连接去重
                    agg_df = sub_df.groupby(p_keys + [item_col])[act_col].apply(lambda x: " | ".join(x.astype(str).unique())).reset_index()
                    agg_df['数值'] = agg_df[act_col]

                agg_df['指标分类'] = " ⭐ " + agg_df[item_col].astype(str)
                agg_df['时长'] = ""
                agg_df['排序层级'] = order
                
                standard_dfs.append(agg_df[p_keys + ['指标分类', '时长', '数值', '排序层级']])

    main_df = pd.concat(standard_dfs, ignore_index=True) if standard_dfs else None
    if main_df is not None:
        sort_cols = [k for k in JOIN_KEYS if k in main_df.columns] + ['排序层级']
        main_df = main_df.sort_values(by=sort_cols).reset_index(drop=True)
        # 暂时保留 '排序层级' 供前端透视排序使用

    return main_df


def load_app_detail_raw():
    """加载应用明细原始数据（不聚合，供钻取页面使用）"""
    cfg = next((c for c in TABLE_CONFIGS if c.get('type') == 'app_detail'), None)
    if not cfg:
        return None
    
    all_files = glob.glob(os.path.join(DATA_DIR, "*.xls*"))
    dfs = []
    for fp in all_files:
        if cfg['keyword'].lower() in os.path.basename(fp).lower():
            try:
                df = pd.read_excel(fp)
                df.columns = df.columns.str.strip()
                if 'data' in df.columns:
                    df['data'] = pd.to_datetime(df['data']).dt.strftime('%Y-%m-%d')
                dfs.append(df)
            except Exception:
                continue
    if dfs:
        return pd.concat(dfs, ignore_index=True)
    return None


with st.spinner("加载后台多维大屏数据..."):
    main_df = load_data()
    app_detail_df = load_app_detail_raw()

# =====================================================================
# 📊 上半屏：宏观诊断与静态数值看板 (主表)
# =====================================================================
if main_df is not None and not main_df.empty:
    # ==========================================
    # 🎯 全局控制中枢 (侧边栏)
    # ==========================================
    with st.sidebar:
        st.markdown("## 🧭 系统导航")
        # --- 菜单切换核心 ---
        menu = st.radio("请选择控制台模块", ["📊 全景数据大盘", "📦 报表导出中心", "📱 应用钻取分析"])
        
    # --- 全局数据漏斗（统一渲染，路由分发前执行） ---
    date_series = pd.to_datetime(main_df["data"].dropna())
    min_date = date_series.min().date()
    max_date = date_series.max().date()

    # 根据当前菜单动态切换页面标题
    if menu == "📊 全景数据大盘":
        st.markdown('<div class="title-box"><h2>🌌 智能数据中台：全景状态追溯诊断大屏</h2></div>', unsafe_allow_html=True)
    elif menu == "📦 报表导出中心":
        st.markdown('<div class="title-box" style="background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%);"><h2>📦 智能数据中台：高级报表批量导出车间</h2></div>', unsafe_allow_html=True)

    # 漏斗 UI（仅渲染一次，固定 key，不随菜单变化）
    with st.expander("🔍 数据漏斗 (筛选范围)", expanded=True):
        with st.form(key="global_filter_form"):
            col_f1, col_f2, col_f3 = st.columns([3, 3, 2])
            with col_f1:
                selected_id = st.text_input("🎯 设备 ID (模糊搜索，留空查全量)")
            with col_f2:
                selected_date_range = st.date_input("📅 筛选日期区间", value=(min_date, max_date), min_value=min_date, max_value=max_date)
            with col_f3:
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                st.form_submit_button("应用筛选条件", use_container_width=True)

    # 执行过滤逻辑
    filtered_df = main_df.copy()
    if selected_id.strip():
        filtered_df = filtered_df[filtered_df["id"].astype(str).str.contains(selected_id.strip(), case=False, na=False)]
    if selected_date_range:
        if len(selected_date_range) == 2:
            s_d, e_d = selected_date_range
            filtered_df = filtered_df[
                (pd.to_datetime(filtered_df["data"]).dt.date >= s_d) &
                (pd.to_datetime(filtered_df["data"]).dt.date <= e_d)
            ]
        elif len(selected_date_range) == 1:
            s_d = selected_date_range[0]
            filtered_df = filtered_df[pd.to_datetime(filtered_df["data"]).dt.date == s_d]

    # 公共工具函数
    def hide_repeated(df_target, col_name):
        if col_name in df_target.columns:
            s = df_target[col_name].copy()
            s[s == s.shift()] = ""
            df_target[col_name] = s
        return df_target

    # ==========================================
    # 🗂️ 路由分发区 (主舞台)
    # ==========================================
    if menu == "📊 全景数据大盘":
        # --- 顶层 Metrics 数据板 ---
        total_devices = filtered_df['id'].nunique()
        days_span = (s_d - s_d).days + 1 if len(selected_date_range) == 1 else ((selected_date_range[1] - selected_date_range[0]).days + 1 if len(selected_date_range) == 2 else 1)
        total_records = len(filtered_df)

        col1, col2, col3 = st.columns(3)
        col1.metric("🖥️ 当前检出设备", f"{total_devices} 台")
        col2.metric("📅 追踪时间跨度", f"{days_span} 天")
        col3.metric("📊 追溯记录总数", f"{total_records} 条")

        st.markdown("<hr style='margin: 15px 0; border: none; border-top: 1px dashed #ccc;'/>", unsafe_allow_html=True)
        st.markdown("### 🗂️ 全局状态追溯卷宗")

        ui_display_df = filtered_df.drop(columns=['排序层级'], errors='ignore')
        ui_display_df = hide_repeated(ui_display_df, 'id')

        st.dataframe(ui_display_df, use_container_width=True, height=700, hide_index=True)

    elif menu == "📦 报表导出中心":
        st.info("💡 下方导出策略已自动继承页面上方的【数据漏斗】范围。您可以直接全量生成，或开启微操对局部设备进行干预。")

        
        available_ids = sorted(filtered_df['id'].dropna().unique().tolist())
        export_date_config = {}
        
        if available_ids:
            # 2. 使用全宽度的卡片来放置这些内容
            with st.container():
                st.markdown("#### ⚙️ 高级配置 (选配)")
                enable_micro_config = st.checkbox("🚀 开启设备级独立日期选配 (干预特定设备的导出日期)")
                
                if enable_micro_config:
                    st.caption("为避免海量设备造成配置面板冗长，请首先挑选需要干预的设备：")
                    target_micro_devices = st.multiselect("🎯 检索并选中需要特别配置的设备", available_ids)
                    
                    if target_micro_devices:
                        st.markdown("##### 👇 分设备日期配置台")
                        # 动态生成多列布局以平铺设备的配置框，极大提高空间利用率
                        cols = st.columns(min(len(target_micro_devices), 3) or 1)
                        for idx, dev_id in enumerate(target_micro_devices):
                            dev_dates = sorted(filtered_df[filtered_df['id'] == dev_id]['data'].unique().tolist())
                            with cols[idx % 3]:
                                export_date_config[dev_id] = st.multiselect(
                                    f"👉 {dev_id}",
                                    options=dev_dates,
                                    default=dev_dates,
                                    key=f"export_{dev_id}"
                                )
                    
                    # 未干预的设备默认导出所有可见数据
                    for dev_id in available_ids:
                        if dev_id not in target_micro_devices:
                            export_date_config[dev_id] = sorted(filtered_df[filtered_df['id'] == dev_id]['data'].unique().tolist())
                else:
                    # 关闭高级微操时，默认全部导出
                    for dev_id in available_ids:
                        export_date_config[dev_id] = sorted(filtered_df[filtered_df['id'] == dev_id]['data'].unique().tolist())
                        
            # 3. 导出生成逻辑
            output = io.BytesIO()
            has_data = False
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                sheet_name = '透视诊断结果'
                current_row = 0
                
                if export_date_config:
                    for dev_id, sel_dates in export_date_config.items():
                        if sel_dates:
                            chunk = filtered_df[(filtered_df["id"] == dev_id) & (filtered_df["data"].isin(sel_dates))]
                            if not chunk.empty and 'data' in chunk.columns:
                                has_data = True
                                
                                pivot_df = chunk.pivot_table(index=['id', '排序层级', '指标分类'], columns='data', values=['时长', '数值'], aggfunc='first')
                                new_cols = [f"{c[1]} {c[0]}" for c in pivot_df.columns]
                                pivot_df.columns = new_cols
                                
                                def col_sort_key(c):
                                    if ' 时长' in c: return c.replace(' 时长', '_0_时长')
                                    elif ' 数值' in c: return c.replace(' 数值', '_1_数值')
                                    return c
                                    
                                pivot_df = pivot_df.reindex(sorted(pivot_df.columns, key=col_sort_key), axis=1).reset_index()
                                pivot_df = pivot_df.sort_values(by=['id', '排序层级']).reset_index(drop=True)
                                pivot_df.drop(columns=['排序层级'], inplace=True, errors='ignore')
                                
                                def hide_rep(df_target, col_name):
                                    if col_name in df_target.columns:
                                        s = df_target[col_name].copy()
                                        s[s == s.shift()] = ""
                                        df_target[col_name] = s
                                    return df_target
                                pivot_df = hide_rep(pivot_df, 'id')
                                
                                pivot_df.to_excel(writer, index=False, sheet_name=sheet_name, startrow=current_row)
                                current_row += len(pivot_df) + 2
                                
                if not has_data:
                    pd.DataFrame({"提示": ["暂无数据或未在配置中勾选日期"]}).to_excel(writer, index=False, sheet_name='空数据')
                else:
                    worksheet = writer.sheets[sheet_name]
                    worksheet.column_dimensions['A'].width = 18
                    worksheet.column_dimensions['B'].width = 35
                    for i in range(3, min(26, worksheet.max_column + 1)):
                        worksheet.column_dimensions[chr(64 + i)].width = 18
                    
            excel_data = output.getvalue()
            
            st.markdown("<hr style='margin: 30px 0; border: none; border-top: 1px solid #eee;'/>", unsafe_allow_html=True)
            col_btn_center = st.columns([1, 2, 1])
            with col_btn_center[1]:
                st.download_button(
                    label="📥 立即生成并下载级联式宽表 (Excel)",
                    data=excel_data,
                    file_name="诊断结果_断层级联版.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
        else:
            st.warning("⚠️ 全局漏斗未命中任何数据，请放宽搜索条件。")

    elif menu == "📱 应用钻取分析":
        st.markdown('<div class="title-box" style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);"><h2>📱 应用运行明细钻取分析</h2></div>', unsafe_allow_html=True)
        
        if app_detail_df is not None and not app_detail_df.empty:
            cfg = next((c for c in TABLE_CONFIGS if c.get('type') == 'app_detail'), {})
            app_col = cfg.get('app_col', '应用名称')
            dur_col = cfg.get('duration_col', '运行时长')
            pwr_col = cfg.get('power_col', '功耗值')
            
            # 筛选控制条
            st.markdown("#### 🎯 钻取条件")
            ctl1, ctl2, ctl3, ctl4 = st.columns([2, 2, 2, 2])
            
            with ctl1:
                avail_ids = sorted(app_detail_df['id'].unique().tolist())
                sel_dev = st.selectbox("🖥️ 选择设备", avail_ids)
            
            dev_df = app_detail_df[app_detail_df['id'] == sel_dev]
            
            with ctl2:
                avail_dates = sorted(dev_df['data'].unique().tolist())
                sel_date = st.selectbox("📅 选择日期", avail_dates)
            
            date_df = dev_df[dev_df['data'] == sel_date]
            
            with ctl3:
                avail_apps = sorted(date_df[app_col].unique().tolist())
                sel_app = st.selectbox("📱 选择应用", ["全部应用"] + avail_apps)
            
            with ctl4:
                min_duration = st.number_input("⏱ 时长阈值 (秒)，仅展示≥此值的记录", min_value=0, value=0, step=10)
            
            # 执行钻取过滤
            detail = date_df.copy()
            if sel_app != "全部应用":
                detail = detail[detail[app_col] == sel_app]
            if min_duration > 0 and dur_col in detail.columns:
                detail[dur_col] = pd.to_numeric(detail[dur_col], errors='coerce').fillna(0)
                detail = detail[detail[dur_col] >= min_duration]
            
            st.markdown("<hr style='margin: 15px 0; border: none; border-top: 1px dashed #ccc;'/>", unsafe_allow_html=True)
            
            if detail.empty:
                st.warning("⚠️ 当前条件下无匹配记录，请调低时长阈值或更换筛选。")
            else:
                # 汇总指标
                total_runs = len(detail)
                total_dur = pd.to_numeric(detail[dur_col], errors='coerce').sum()
                total_pwr = pd.to_numeric(detail[pwr_col], errors='coerce').sum()
                avg_dur = total_dur / total_runs if total_runs > 0 else 0
                
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("🔁 运行次数", f"{total_runs} 次")
                m2.metric("⏱ 总时长", f"{total_dur:.0f} 秒")
                m3.metric("⚡ 总功耗", f"{total_pwr:.0f} mA")
                m4.metric("📊 平均时长", f"{avg_dur:.1f} 秒/次")
                
                st.markdown("<hr style='margin: 15px 0; border: none; border-top: 1px dashed #ccc;'/>", unsafe_allow_html=True)
                
                # 可视化区域
                chart_df = detail.copy()
                chart_df[dur_col] = pd.to_numeric(chart_df[dur_col], errors='coerce')
                chart_df[pwr_col] = pd.to_numeric(chart_df[pwr_col], errors='coerce')
                chart_df = chart_df.reset_index(drop=True)
                
                # === 第一行：环形占比 + 气泡散点 ===
                row1_c1, row1_c2 = st.columns(2)
                
                with row1_c1:
                    # 按应用聚合总时长，做环形图
                    agg_app = chart_df.groupby(app_col).agg(
                        总时长=(dur_col, 'sum'),
                        总功耗=(pwr_col, 'sum'),
                        运行次数=(dur_col, 'count')
                    ).reset_index()
                    
                    fig_donut = px.pie(
                        agg_app, names=app_col, values='总时长',
                        title="🍩 各应用运行时长占比",
                        hole=0.45,
                        color_discrete_sequence=px.colors.qualitative.Set2
                    )
                    fig_donut.update_traces(textposition='inside', textinfo='label+percent')
                    fig_donut.update_layout(height=380, showlegend=False)
                    st.plotly_chart(fig_donut, use_container_width=True)
                
                with row1_c2:
                    # 散点图：时长 vs 功耗，气泡大小 = 功耗
                    fig_scatter = px.scatter(
                        chart_df, x=dur_col, y=pwr_col,
                        color=app_col, size=pwr_col,
                        title="🫧 时长 × 功耗关联分析 (气泡越大功耗越高)",
                        labels={dur_col: "运行时长 (秒)", pwr_col: "功耗 (mA)"},
                        color_discrete_sequence=px.colors.qualitative.Bold,
                        size_max=40
                    )
                    fig_scatter.update_layout(height=380)
                    st.plotly_chart(fig_scatter, use_container_width=True)
                
                # === 第二行：分应用堆叠柱状 + Treemap 全景 ===
                row2_c1, row2_c2 = st.columns(2)
                
                with row2_c1:
                    # 按应用分组的堆叠柱状图
                    chart_df['运行编号'] = chart_df.groupby(app_col).cumcount() + 1
                    chart_df['标签'] = chart_df[app_col] + " #" + chart_df['运行编号'].astype(str)
                    
                    fig_bar = px.bar(
                        chart_df, x=app_col, y=dur_col, color='标签',
                        title="📊 分应用运行时长堆叠",
                        labels={dur_col: "时长 (秒)", app_col: "应用"},
                        color_discrete_sequence=px.colors.qualitative.Pastel
                    )
                    fig_bar.update_layout(height=380, showlegend=True, barmode='stack')
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                with row2_c2:
                    # Treemap 全景：面积 = 时长，颜色 = 功耗
                    fig_tree = px.treemap(
                        chart_df, path=[app_col], values=dur_col,
                        color=pwr_col,
                        title="🗺️ 应用消耗全景 (面积=时长, 颜色深浅=功耗)",
                        color_continuous_scale='YlOrRd'
                    )
                    fig_tree.update_layout(height=380)
                    st.plotly_chart(fig_tree, use_container_width=True)
                
                # === 底部：明细数据表 ===
                st.markdown("### 📝 运行明细记录")
                display_cols = [c for c in ['id', 'data', app_col, dur_col, pwr_col] if c in detail.columns]
                st.dataframe(detail[display_cols].reset_index(drop=True), use_container_width=True, height=400, hide_index=True)
        else:
            st.warning("🚨 未检测到应用明细数据，请确认 data_source 目录中包含带 'app_detail' 关键字的 Excel 文件。")

else:
    st.warning("🚨 后台未检测到任何指标配置及对应的关联数据，请检查 data_source 目录。")
