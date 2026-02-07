import streamlit as st
import pandas as pd
import io
import random

# --- 設定頁面 (Page Config) ---
st.set_page_config(page_title="Taiwan CDM 戰情室 Pro", layout="wide")

# --- 核心資料庫：台灣資安廠商清單 (作為預覽與 AI 推薦基礎) ---
solutions_db = {
    ("裝置", "識別"): ["一休資訊", "台達電子", "思邦科技", "瑞恩資訊"],
    ("裝置", "保護"): ["三甲科技", "安碁資訊", "勤業眾信", "趨勢科技"],
    ("裝置", "偵測"): ["元盾資安", "伊雲谷", "動力安全"],
    ("裝置", "應變"): ["中芯數據", "元盾資安", "安碁資訊"],
    ("裝置", "復原"): ["扇原科技", "肇真數位"],
    
    ("應用程式", "識別"): ["又碩電腦", "元盾資安", "系微"],
    ("應用程式", "保護"): ["三甲科技", "台眾電腦", "安侯企管"],
    ("應用程式", "偵測"): ["安碁資訊", "鼎原科技"],
    ("應用程式", "應變"): ["中芯數據", "宏基資訊"],
    ("應用程式", "復原"): ["安碁資訊"],

    ("網路", "識別"): ["三甲科技", "安碁資訊", "承映資訊"],
    ("網路", "保護"): ["一休資訊", "台眾電腦", "池安量子"],
    ("網路", "偵測"): ["中飛科技", "思邦科技", "雲智維"],
    ("網路", "應變"): ["三甲科技", "元盾資安"],
    ("網路", "復原"): ["如梭世代", "動力安全"],

    ("資料", "識別"): ["台眾電腦", "安碁資訊", "中華電信"],
    ("資料", "保護"): ["三甲科技", "台灣信威", "帝璽智慧"],
    ("資料", "偵測"): ["安碁資訊"],
    ("資料", "應變"): ["三甲科技", "元盾資安"],
    ("資料", "復原"): ["三甲科技", "云碩科技"],

    ("使用者", "識別"): ["一休資訊", "帝濶智慧", "全球系統"],
    ("使用者", "保護"): ["又碩電腦", "全域科技", "希臘智慧"],
    ("使用者", "偵測"): ["伊雲谷"],
    ("使用者", "應變"): ["三甲科技", "肇真數位"],
    ("使用者", "復原"): ["思邦科技"],
}

# --- 初始化 Session State ---
# 注意：這裡將 '皇冠寶石' 改為 '關鍵資產'
if 'assets' not in st.session_state:
    st.session_state.assets = pd.DataFrame(columns=["資產名稱", "類別", "關鍵資產"])
if 'assessments' not in st.session_state:
    st.session_state.assessments = {}
if 'current_page' not in st.session_state:
    st.session_state.current_page = "1. 資產盤點"

# --- 側邊欄導航 ---
st.sidebar.title("🛡️ Taiwan CDM Pro")
pages = ["1. 資產盤點", "2. 防禦診斷", "3. 風險戰情室"]
page_selection = st.sidebar.radio("導航", pages, index=pages.index(st.session_state.current_page))

if page_selection != st.session_state.current_page:
    st.session_state.current_page = page_selection
    st.rerun()

# --- 輔助函式：生成 Excel (修正欄位名稱) ---
def create_template_excel(mode="empty"):
    columns = ["資產名稱", "類別", "關鍵資產"]
    
    if mode == "example":
        # === 生成 50 筆測試資料 ===
        categories = ["裝置", "應用程式", "網路", "資料", "使用者"]
        data = []
        
        for cat in categories:
            for i in range(1, 11): 
                # 20% 機率為關鍵資產
                is_critical = "是" if random.random() < 0.2 else "否"
                
                if cat == "裝置":
                    name = f"員工筆電_{random.randint(100,999)}" if i > 3 else f"核心伺服器_{i:02d}"
                elif cat == "應用程式":
                    name = f"HR系統_{i}" if i % 2 == 0 else f"ERP模組_{i}"
                elif cat == "網路":
                    name = f"防火牆_{i:02d}" if i < 3 else f"辦公區Switch_{i:02d}"
                elif cat == "資料":
                    name = f"客戶個資DB_{i}" if is_critical == "是" else f"日常備份_{i}"
                else: 
                    name = f"行政人員_{i:02d}" if i > 2 else f"系統管理員_{i:02d}"
                    
                data.append([name, cat, is_critical])
        
        df = pd.DataFrame(data, columns=columns)
    else:
        df = pd.DataFrame(columns=columns)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='資產清單')
    buffer.seek(0)
    return buffer

# --- 核心邏輯函數：計算 CDM 格子狀態 (修正邏輯) ---
def calculate_cell_status(category, function):
    df = st.session_state.assets
    if df.empty: return "no_asset", 0, []
    
    related_assets = df[df['類別'] == category]
    if related_assets.empty: return "no_asset", 0, []

    scores = []
    has_critical_risk = False  # 變數名稱更新
    details = []

    for index, row in related_assets.iterrows():
        asset_name = row['資產名稱']
        is_critical = row['關鍵資產'] # 讀取新欄位
        key = (asset_name, function)
        
        score = st.session_state.assessments.get(key, 0)
        
        if score > 0:
            scores.append(score)
            details.append(f"{asset_name}: Tier {score}")
            # 如果是關鍵資產，且防護等級低於 Tier 3 (即 < 50%)，視為高風險
            if is_critical and score < 3:
                has_critical_risk = True
    
    if not scores: return "not_assessed", 0, []

    if has_critical_risk: return "critical_risk", 1, details # 狀態名稱更新
    
    avg_score = sum(scores) / len(scores)
    # 根據四原則 (百分比概念)
    if avg_score < 1.5: return "tier-1", 1, details
    elif avg_score < 2.5: return "tier-2", 2, details
    elif avg_score < 3.5: return "tier-3", 3, details
    else: return "tier-4", 4, details

# ==========================================
# 頁面 1: 資產盤點 (Inventory)
# ==========================================
if st.session_state.current_page == "1. 資產盤點":
    st.header("📍 步驟一：建立戰場地圖 (Inventory)")
    
    # --- 區塊 A: 批次匯入 / 下載 ---
    with st.expander("📤 批次匯入 / 下載範本 (Excel)", expanded=True):
        col_dl, col_up = st.columns([1, 1])
        
        with col_dl:
            st.markdown("#### 1. 取得格式")
            b1, b2 = st.columns(2)
            with b1:
                st.download_button(
                    label="📄 空白範本",
                    data=create_template_excel(mode="empty"),
                    file_name="CDM_空白盤點表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with b2:
                st.download_button(
                    label="🎲 50筆範例",
                    data=create_template_excel(mode="example"),
                    file_name="CDM_範例資料(50筆).xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with col_up:
            st.markdown("#### 2. 上傳盤點表")
            uploaded_file = st.file_uploader("支援 .xlsx，上傳後**覆蓋**現有資料", type=["xlsx"])
            
            if uploaded_file is not None:
                try:
                    df_new = pd.read_excel(uploaded_file)
                    
                    # 檢查欄位 (支援舊版 '皇冠寶石' 自動轉移，或新版 '關鍵資產')
                    if "皇冠寶石" in df_new.columns:
                        df_new.rename(columns={"皇冠寶石": "關鍵資產"}, inplace=True)
                    
                    required_cols = {"資產名稱", "類別", "關鍵資產"}
                    
                    if not required_cols.issubset(df_new.columns):
                        st.error(f"❌ 格式錯誤！缺少欄位：{required_cols}")
                    else:
                        if st.button("✅ 確認匯入 (覆蓋)", type="primary"):
                            # 資料清洗
                            def parse_bool(val):
                                val_str = str(val).lower().strip()
                                return val_str in ["yes", "y", "是", "true", "1"]
                            
                            df_new["關鍵資產"] = df_new["關鍵資產"].apply(parse_bool)
                            
                            valid_cats = ["裝置", "應用程式", "網路", "資料", "使用者"]
                            df_new["類別"] = df_new["類別"].apply(lambda x: x if x in valid_cats else "裝置")

                            st.session_state.assets = df_new
                            st.session_state.assessments = {} 
                            st.success("🎉 匯入成功！")
                            st.rerun()
                except Exception as e:
                    st.error(f"讀取失敗：{e}")

    st.divider()

    # --- 區塊 B: 手動新增 (UI調整) ---
    st.subheader("✍️ 手動新增資產")
    with st.container():
        col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
        with col1: asset_name = st.text_input("資產名稱", placeholder="例: 核心資料庫")
        with col2: asset_type = st.selectbox("類別", ["裝置", "應用程式", "網路", "資料", "使用者"])
        # 這裡改用您建議的更直觀的描述
        with col3: is_critical = st.checkbox("🔥 設為關鍵資產", help="即 CDM 的『皇冠寶石』，指公司最重要的資產")
        with col4: 
            st.write("") 
            st.write("")
            add_btn = st.button("新增", use_container_width=True)
        
        if add_btn:
            if asset_name:
                current_names = st.session_state.assets['資產名稱'].values if not st.session_state.assets.empty else []
                if asset_name not in current_names:
                    # 寫入 '關鍵資產'
                    new_row = {"資產名稱": asset_name, "類別": asset_type, "關鍵資產": is_critical}
                    st.session_state.assets = pd.concat([st.session_state.assets, pd.DataFrame([new_row])], ignore_index=True)
                    st.success(f"已新增: {asset_name}")
                else:
                    st.warning("資產名稱重複！")
            else:
                st.error("請輸入名稱")

    # --- 區塊 C: 清單顯示 ---
    if not st.session_state.assets.empty:
        st.subheader(f"📋 資產清單 ({len(st.session_state.assets)} 筆)")
        
        # 顯示時將 True 高亮為金色
        def highlight_critical(val): return 'background-color: #ffd700; color: black; font-weight: bold' if val else ''
        
        st.dataframe(
            st.session_state.assets.style.applymap(highlight_critical, subset=['關鍵資產']), 
            use_container_width=True,
            hide_index=True
        )
        
        if st.button("🗑️ 清空所有資產"):
            st.session_state.assets = pd.DataFrame(columns=["資產名稱", "類別", "關鍵資產"])
            st.session_state.assessments = {}
            st.rerun()
    else:
        st.info("👈 清單為空，請匯入或新增。")

    st.divider()
    if st.button("下一步：防禦診斷 👉", use_container_width=True):
        st.session_state.current_page = "2. 防禦診斷"
        st.rerun()

# ==========================================
# 頁面 2: 防禦診斷 (Assessment)
# ==========================================
elif st.session_state.current_page == "2. 防禦診斷":
    st.header("🩺 步驟二：防禦成熟度診斷")
    
    with st.expander("📊 評分標準說明 (依防護覆蓋率)", expanded=True):
        st.markdown("""
        * **0 (N/A)**: 不適用。
        * **Tier 1 (< 25%)**: 🔴 **低度防護**。無防護或僅有少數安裝，且無SOP。
        * **Tier 2 (25-50%)**: 🟡 **基礎防護**。有設備但流程未標準化，依賴人員記憶。
        * **Tier 3 (50-75%)**: 🟢 **標準防護**。有白紙黑字 SOP，且 75% 資產已受控。
        * **Tier 4 (> 75%)**: 🏆 **進階防護**。高度自動化偵測與回應。
        """)

    target_category = st.selectbox("請選擇要評估的類別：", ["裝置", "應用程式", "網路", "資料", "使用者"])
    
    assets_in_cat = st.session_state.assets[st.session_state.assets['類別'] == target_category]
    
    if assets_in_cat.empty:
        st.warning(f"⚠️ 尚未建立「{target_category}」類別的資產，請回上一步新增。")
    else:
        st.info(f"正在評估 {len(assets_in_cat)} 項資產。")
        
        tabs = st.tabs(["識別 (ID)", "保護 (PR)", "偵測 (DE)", "應變 (RS)", "復原 (RC)"])
        functions = ["識別", "保護", "偵測", "應變", "復原"]
        
        for i, func in enumerate(functions):
            with tabs[i]:
                # 進度條
                total_items = len(assets_in_cat)
                assessed_count = 0
                for _, row in assets_in_cat.iterrows():
                     if st.session_state.assessments.get((row['資產名稱'], func), 0) > 0:
                         assessed_count += 1
                
                prog = assessed_count / total_items if total_items > 0 else 0
                st.progress(prog, text=f"完成度: {int(prog*100)}%")

                for idx, row in assets_in_cat.iterrows():
                    asset = row['資產名稱']
                    is_critical = row['關鍵資產']
                    # 圖示改為火或是鑽石
                    critical_label = "🔥" if is_critical else ""
                    
                    key = (asset, func)
                    current_val = st.session_state.assessments.get(key, 0)
                    
                    with st.container():
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            st.markdown(f"**{asset}** {critical_label}")
                            if is_critical:
                                st.caption("關鍵資產")
                        with c2:
                            # 使用 Slider 讓百分比更直觀
                            score = st.select_slider(
                                f"防護覆蓋率 ({asset}-{func})",
                                options=[0, 1, 2, 3, 4],
                                value=current_val,
                                format_func=lambda x: {
                                    0: "⚪ N/A",
                                    1: "🔴 <25% (低)",
                                    2: "🟡 25-50% (中)",
                                    3: "🟢 50-75% (高)",
                                    4: "🏆 >75% (全)"
                                }[x],
                                key=f"slider_{asset}_{func}",
                                label_visibility="collapsed"
                            )
                            if score != current_val:
                                st.session_state.assessments[key] = score
                        st.divider()

    col_prev, col_next = st.columns(2)
    with col_prev:
        if st.button("👈 上一步", use_container_width=True):
            st.session_state.current_page = "1. 資產盤點"
            st.rerun()
    with col_next:
        if st.button("下一步：進入戰情室 👉", use_container_width=True):
            st.session_state.current_page = "3. 風險戰情室"
            st.rerun()

# ==========================================
# 頁面 3: 風險戰情室 (Dashboard)
# ==========================================
elif st.session_state.current_page == "3. 風險戰情室":
    st.header("📊 步驟三：CDM 風險戰情室")
    
    categories = ["裝置", "應用程式", "網路", "資料", "使用者"]
    functions = ["識別", "保護", "偵測", "應變", "復原"]
    recommendation_list = []

    # --- HTML 矩陣 (CSS 優化避免 br) ---
    html_code = """
    <style>
        table {width: 100%; border-collapse: separate; border-spacing: 4px;}
        th {background-color: #262730; color: white; padding: 10px; font-size: 0.9em;}
        td {
            padding: 8px; height: 70px; text-align: center; vertical-align: middle;
            border-radius: 6px; font-weight: bold; font-size: 0.9em; color: black;
            box-shadow: 1px 1px 2px rgba(0,0,0,0.1);
        }
        .cat-head {background-color: #555; color: white; width: 15%;}
        
        .s-no-asset {background-color: #f0f2f6; color: #ccc; border: 1px dashed #ddd;}
        .s-pending {background-color: #ffffff; color: #888; border: 1px solid #ddd;}
        
        /* 關鍵資產風險樣式 */
        .s-critical-risk {
            background-color: #ff4b4b; color: white; 
            border: 3px solid #ffd700; 
            animation: pulse 2s infinite;
        }
        
        .s-t1 {background-color: #ffcccc; border: 1px solid red;}
        .s-t2 {background-color: #fff3cd; border: 1px solid orange;}
        .s-t3 {background-color: #d1e7dd; border: 1px solid green;}
        .s-t4 {background-color: #cff4fc; border: 2px solid #0dcaf0;}
        
        @keyframes pulse { 0% {box-shadow: 0 0 0 0 rgba(255, 75, 75, 0.7);} 70% {box-shadow: 0 0 0 10px rgba(255, 75, 75, 0);} 100% {box-shadow: 0 0 0 0 rgba(255, 75, 75, 0);} }
    </style>
    <table>
        <tr><th>CDM</th>
    """
    for f in functions: html_code += f"<th>{f}</th>"
    html_code += "</tr>"

    for cat in categories:
        html_code += f"<tr><td class='cat-head'>{cat}</td>"
        for func in functions:
            status, tier, details = calculate_cell_status(cat, func)
            
            # 這裡把 critical_risk 加入推薦清單
            if status in ["critical_risk", "tier-1", "tier-2"]:
                recommendation_list.append((cat, func, status))

            css_class = ""
            display_text = ""
            
            if status == "no_asset":
                css_class = "s-no-asset"
                display_text = "."
            elif status == "not_assessed":
                css_class = "s-pending"
                display_text = "?"
            elif status == "critical_risk":
                css_class = "s-critical-risk"
                display_text = "🔥 關鍵<br>風險" # 這裡手動換行比較美觀，但 CSS 支援
            else:
                css_class = f"s-t{tier}"
                display_text = f"Tier {tier}"
            
            # 為了避免 title 屬性導致的亂碼，我們不放 title，或是確保 details 是純文字
            clean_details = "&#10;".join(details) # HTML tooltip 換行符
            html_code += f"<td class='{css_class}' title='{clean_details}'>{display_text}</td>"
        html_code += "</tr>"
    html_code += "</table>"
    
    st.markdown(html_code, unsafe_allow_html=True)

    # --- 處方籤 ---
    st.divider()
    st.subheader("💊 智慧處方籤 (AI 推薦 x SecPaaS)")
    
    SECPAAS_URL = "https://secpaas.org.tw/W_SecDocProduct"
    
    if recommendation_list:
        st.write(f"共偵測到 **{len(recommendation_list)}** 個弱點：")
        
        for cat, func, status in recommendation_list:
            if status == "critical_risk":
                label = "🔥 關鍵資產告急"
                desc = "重要資產防護低於標準 (Tier 3)，風險極高！"
            elif status == "tier-1":
                label = "🔴 Tier 1 缺口"
                desc = "防護覆蓋率 < 25%，形同裸奔。"
            else:
                label = "🟡 Tier 2 強化"
                desc = "防護覆蓋率 < 50%，需建立標準流程 (SOP)。"
            
            with st.expander(f"{label}：[{cat} - {func}]", expanded=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    vendors = solutions_db.get((cat, func), [])
                    vendor_txt = "、".join(vendors[:4]) + ("..." if len(vendors)>4 else "") if vendors else "請查詢 SecPaaS"
                    st.markdown(f"**診斷：** {desc}")
                    st.markdown(f"👀 **建議廠商：** {vendor_txt}")
                with c2:
                    st.write("")
                    st.link_button("🔍 找解方", url=SECPAAS_URL)
    else:
        if st.session_state.assets.empty:
            st.warning("⚠️ 無資產資料。")
        else:
            st.success("🎉 恭喜！無重大風險紅燈。")
            st.link_button("前往 SecPaaS 商城", SECPAAS_URL)

    st.write("")
    if st.button("🔄 重新盤點", use_container_width=True):
        st.session_state.current_page = "1. 資產盤點"
        st.rerun()
