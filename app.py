import streamlit as st
import pandas as pd
import io
import random

# --- 設定頁面 (Page Config) ---
st.set_page_config(page_title="Taiwan CDM 戰情室 Pro", layout="wide")

# --- 核心資料庫：台灣資安廠商清單 (作為預覽與 AI 推薦基礎) ---
solutions_db = {
    ("裝置", "識別"): ["一休資訊", "台達電子", "思邦科技", "瑞恩資訊", "中芯數據", "中華龍網"],
    ("裝置", "保護"): ["三甲科技", "安碁資訊", "勤業眾信", "趨勢科技", "奧義智慧"],
    ("裝置", "偵測"): ["元盾資安", "伊雲谷", "動力安全", "誠雲科技"],
    ("裝置", "應變"): ["中芯數據", "元盾資安", "安碁資訊"],
    ("裝置", "復原"): ["扇原科技", "肇真數位"],
    
    ("應用程式", "識別"): ["又碩電腦", "元盾資安", "系微", "保華資安"],
    ("應用程式", "保護"): ["三甲科技", "台眾電腦", "安侯企管", "瑞恩資訊"],
    ("應用程式", "偵測"): ["安碁資訊", "鼎原科技"],
    ("應用程式", "應變"): ["中芯數據", "宏基資訊", "動力安全"],
    ("應用程式", "復原"): ["安碁資訊"],

    ("網路", "識別"): ["三甲科技", "安碁資訊", "承映資訊"],
    ("網路", "保護"): ["一休資訊", "台眾電腦", "池安量子", "威碩系統"],
    ("網路", "偵測"): ["中飛科技", "思邦科技", "雲智維"],
    ("網路", "應變"): ["三甲科技", "元盾資安", "如梭世代"],
    ("網路", "復原"): ["如梭世代", "動力安全"],

    ("資料", "識別"): ["台眾電腦", "安碁資訊", "中華電信"],
    ("資料", "保護"): ["三甲科技", "台灣信威", "帝璽智慧"],
    ("資料", "偵測"): ["安碁資訊"],
    ("資料", "應變"): ["三甲科技", "元盾資安"],
    ("資料", "復原"): ["三甲科技", "云碩科技", "華碩雲端"],

    ("使用者", "識別"): ["一休資訊", "帝濶智慧", "全球系統"],
    ("使用者", "保護"): ["又碩電腦", "全域科技", "希臘智慧"],
    ("使用者", "偵測"): ["伊雲谷"],
    ("使用者", "應變"): ["三甲科技", "肇真數位"],
    ("使用者", "復原"): ["思邦科技"],
}

# --- 初始化 Session State ---
if 'assets' not in st.session_state:
    st.session_state.assets = pd.DataFrame(columns=["資產名稱", "類別", "皇冠寶石"])
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

# --- 輔助函式：生成 Excel (支援空白或範例) ---
def create_template_excel(mode="empty"):
    columns = ["資產名稱", "類別", "皇冠寶石"]
    
    if mode == "example":
        # === 生成 50 筆測試資料 ===
        categories = ["裝置", "應用程式", "網路", "資料", "使用者"]
        data = []
        
        for cat in categories:
            for i in range(1, 11): # 每類產生 10 筆，共 50 筆
                is_crown = "是" if random.random() < 0.2 else "否" # 20% 機率為皇冠
                
                # 產生擬真的資產名稱
                if cat == "裝置":
                    name = f"員工筆電_{random.randint(100,999)}" if i > 3 else f"核心伺服器_{i:02d}"
                elif cat == "應用程式":
                    name = f"HR系統_{i}" if i % 2 == 0 else f"ERP模組_{i}"
                elif cat == "網路":
                    name = f"防火牆_{i:02d}" if i < 3 else f"辦公區Switch_{i:02d}"
                elif cat == "資料":
                    name = f"客戶個資DB_{i}" if is_crown == "是" else f"日常備份_{i}"
                else: # 使用者
                    name = f"行政人員_{i:02d}" if i > 2 else f"系統管理員_{i:02d}"
                    
                data.append([name, cat, is_crown])
        
        df = pd.DataFrame(data, columns=columns)
    else:
        # === 生成僅有標題的空白資料 ===
        df = pd.DataFrame(columns=columns)

    # 寫入 Buffer
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='資產清單')
    buffer.seek(0)
    return buffer

# --- 核心邏輯函數：計算 CDM 格子狀態 ---
def calculate_cell_status(category, function):
    df = st.session_state.assets
    if df.empty: return "no_asset", 0, []
    
    related_assets = df[df['類別'] == category]
    if related_assets.empty: return "no_asset", 0, []

    scores = []
    has_crown_risk = False
    details = []

    for index, row in related_assets.iterrows():
        asset_name = row['資產名稱']
        is_crown = row['皇冠寶石']
        key = (asset_name, function)
        
        score = st.session_state.assessments.get(key, 0)
        
        if score > 0:
            scores.append(score)
            details.append(f"{asset_name}: Tier {score}")
            if is_crown and score < 3:
                has_crown_risk = True
    
    if not scores: return "not_assessed", 0, []

    if has_crown_risk: return "crown_risk", 1, details
    
    avg_score = sum(scores) / len(scores)
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
            st.caption("選擇下載：")
            
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
                    
                    # 簡易檢查欄位
                    required_cols = {"資產名稱", "類別", "皇冠寶石"}
                    if not required_cols.issubset(df_new.columns):
                        st.error(f"❌ 格式錯誤！缺少欄位：{required_cols}")
                    else:
                        st.info(f"偵測到 {len(df_new)} 筆資產資料。")
                        
                        if st.button("✅ 確認匯入 (覆蓋)", type="primary"):
                            # 資料清洗：皇冠寶石轉 Boolean
                            def parse_crown(val):
                                val_str = str(val).lower().strip()
                                return val_str in ["yes", "y", "是", "true", "1"]
                            
                            df_new["皇冠寶石"] = df_new["皇冠寶石"].apply(parse_crown)
                            
                            # 資料清洗：類別防呆
                            valid_cats = ["裝置", "應用程式", "網路", "資料", "使用者"]
                            df_new["類別"] = df_new["類別"].apply(lambda x: x if x in valid_cats else "裝置")

                            # 更新 Session State (覆蓋)
                            st.session_state.assets = df_new
                            st.session_state.assessments = {} # 清空舊評分
                            
                            st.success("🎉 匯入成功！資料已更新。")
                            st.rerun()
                            
                except Exception as e:
                    st.error(f"讀取失敗：{e}")

    st.divider()

    # --- 區塊 B: 手動新增 ---
    st.subheader("✍️ 手動新增資產")
    with st.container():
        col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
        with col1: asset_name = st.text_input("資產名稱", placeholder="例: 核心資料庫")
        with col2: asset_type = st.selectbox("類別", ["裝置", "應用程式", "網路", "資料", "使用者"])
        with col3: is_crown = st.checkbox("👑 皇冠寶石?", help="勾選代表此資產極為重要")
        with col4: 
            st.write("") 
            st.write("")
            add_btn = st.button("新增", use_container_width=True)
        
        if add_btn:
            if asset_name:
                current_names = st.session_state.assets['資產名稱'].values if not st.session_state.assets.empty else []
                if asset_name not in current_names:
                    new_row = {"資產名稱": asset_name, "類別": asset_type, "皇冠寶石": is_crown}
                    st.session_state.assets = pd.concat([st.session_state.assets, pd.DataFrame([new_row])], ignore_index=True)
                    st.success(f"已新增: {asset_name}")
                else:
                    st.warning("資產名稱重複！")
            else:
                st.error("請輸入名稱")

    # --- 區塊 C: 清單顯示 ---
    if not st.session_state.assets.empty:
        st.subheader(f"📋 資產清單 ({len(st.session_state.assets)} 筆)")
        
        def highlight_crown(val): return 'background-color: #ffd700; color: black' if val else ''
        
        st.dataframe(
            st.session_state.assets.style.applymap(highlight_crown, subset=['皇冠寶石']), 
            use_container_width=True,
            hide_index=True
        )
        
        if st.button("🗑️ 清空所有資產"):
            st.session_state.assets = pd.DataFrame(columns=["資產名稱", "類別", "皇冠寶石"])
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
    
    # --- 新增：評分指南 Popover (浮動視窗) ---
    with st.popover("📖 打開評分指南 (Cheatsheet)", use_container_width=True):
        st.markdown("### 🔍 NIST CSF 五大功能定義")
        st.markdown("""
        * **識別 (Identify)**：我知道我有什麼資產，以及風險在哪。
        * **保護 (Protect)**：我有裝鎖、裝防毒，讓駭客打不進來。
        * **偵測 (Detect)**：如果駭客進來了，我能馬上發現警告。
        * **應變 (Respond)**：發現被駭，我知道該怎麼止血、通報。
        * **復原 (Recover)**：出事後，我能把資料救回來，恢復營運。
        """)
        st.divider()
        st.markdown("### 📊 成熟度評分標準 (Tier 1-4)")
        st.markdown("""
        | 分數 | 等級 | 定義說明 |
        | :---: | :--- | :--- |
        | **0** | **N/A** | **不適用** (無此功能需求) |
        | **1** | **Tier 1 (被動)** | **有做，但無章法**。<br>通常是發生問題才處理，沒有正式規範。 |
        | **2** | **Tier 2 (部分)** | **有流程，但未全面落實**。<br>知道該怎麼做，但依賴特定人員，沒寫成SOP。 |
        | **3** | **Tier 3 (標準)** | **SOP 標準化**。<br>有白紙黑字規定，且全公司都照著做。 |
        | **4** | **Tier 4 (自動)** | **數據驅動與自動化**。<br>系統能自動阻擋攻擊，並持續優化防禦。 |
        """)

    target_category = st.selectbox("請選擇要評估的類別：", ["裝置", "應用程式", "網路", "資料", "使用者"])
    
    assets_in_cat = st.session_state.assets[st.session_state.assets['類別'] == target_category]
    
    if assets_in_cat.empty:
        st.warning(f"⚠️ 尚未建立「{target_category}」類別的資產，請回上一步新增。")
    else:
        st.info(f"正在評估 {len(assets_in_cat)} 項資產。請點擊上方指南參考評分標準。")
        
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
                    is_crown = row['皇冠寶石']
                    crown_label = "👑" if is_crown else ""
                    
                    key = (asset, func)
                    current_val = st.session_state.assessments.get(key, 0)
                    
                    with st.container():
                        c1, c2 = st.columns([1, 2])
                        with c1:
                            st.markdown(f"**{asset}** {crown_label}")
                        with c2:
                            score = st.radio(
                                f"成熟度 ({asset}-{func})",
                                options=[0, 1, 2, 3, 4],
                                index=current_val,
                                format_func=lambda x: {
                                    0: "⚪ N/A",
                                    1: "🔴 Tier 1",
                                    2: "🟡 Tier 2",
                                    3: "🟢 Tier 3",
                                    4: "🏆 Tier 4"
                                }[x],
                                key=f"radio_{asset}_{func}",
                                horizontal=True,
                                label_visibility="collapsed",
                                help="請參考上方「評分指南」定義分數"
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

    # --- HTML 矩陣 ---
    html_code = """
    <style>
        table {width: 100%; border-collapse: separate; border-spacing: 3px;}
        th {background-color: #333; color: white; padding: 8px; font-size: 0.85em;}
        td {
            padding: 5px; height: 70px; text-align: center; 
            border-radius: 6px; font-weight: bold; font-size: 0.9em; color: black;
            box-shadow: 1px 1px 3px rgba(0,0,0,0.1);
        }
        .cat-head {background-color: #555; color: white; width: 15%;}
        
        .s-no-asset {background-color: #f0f2f6; color: #ccc; border: 1px dashed #ddd;}
        .s-pending {background-color: #ffffff; color: #888; border: 1px solid #ddd;}
        .s-crown-risk {background-color: #ff4b4b; color: white; border: 3px solid #ffd700; animation: pulse 2s infinite;}
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
            
            if status in ["crown_risk", "tier-1", "tier-2"]:
                recommendation_list.append((cat, func, status))

            css_class = ""
            display_text = ""
            
            if status == "no_asset":
                css_class = "s-no-asset"
                display_text = "."
            elif status == "not_assessed":
                css_class = "s-pending"
                display_text = "?"
            elif status == "crown_risk":
                css_class = "s-crown-risk"
                display_text = "⚠️ RISK"
            else:
                css_class = f"s-t{tier}"
                display_text = f"Tier {tier}"
            
            html_code += f"<td class='{css_class}' title='{', '.join(details)}'>{display_text}</td>"
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
            if status == "crown_risk":
                label = "🚨 皇冠風險"
                desc = "關鍵資產防護不足，需立即改善！"
            elif status == "tier-1":
                label = "🔴 Tier 1 缺口"
                desc = "缺乏基礎防禦。"
            else:
                label = "🟡 Tier 2 強化"
                desc = "標準化不足。"
            
            with st.expander(f"{label}：[{cat} - {func}]", expanded=True):
                c1, c2 = st.columns([3, 1])
                with c1:
                    vendors = solutions_db.get((cat, func), [])
                    vendor_txt = "、".join(vendors[:4]) + ("..." if len(vendors)>4 else "") if vendors else "請查詢"
                    st.markdown(f"**診斷：** {desc}")
                    st.markdown(f"👀 **廠商參考：** {vendor_txt}")
                with c2:
                    st.write("")
                    st.link_button("🔍 找廠商", url=SECPAAS_URL)
    else:
        if st.session_state.assets.empty:
            st.warning("⚠️ 無資產資料。")
        else:
            st.success("🎉 無高風險紅燈。")
            st.link_button("前往 SecPaaS", SECPAAS_URL)

    st.write("")
    if st.button("🔄 重新盤點", use_container_width=True):
        st.session_state.current_page = "1. 資產盤點"
        st.rerun()
