import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
import io
import os
import json

# ==========================================
# 網頁基礎與視覺設定
# ==========================================
st.set_page_config(page_title="醫療部病房排班系統", layout="centered")
st.markdown("""
<style>
    .stApp { background-color: #FDFBF7; }
    h1, h2, h3, p, div, span, label { color: #333333; font-family: sans-serif; }
</style>
""", unsafe_allow_html=True)

st.title("醫療部病房值班排班系統")
st.info("📢 **公告：一線醫師開放填寫日期為每月 1 日至 10 日，請於期限內完成登記。**")

# ==========================================
# 核心機制：多人共用資料庫 (JSON)
# ==========================================
PREFS_FILE = 'preferences.json'

def load_preferences():
    if os.path.exists(PREFS_FILE):
        with open(PREFS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_preferences(prefs):
    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
        json.dump(prefs, f, ensure_ascii=False)

# 讀取目前所有醫師已送出的意願
all_prefs = load_preferences()

# ==========================================
# 步驟一：雲端後台自動讀取「醫師總表」
# ==========================================
file_path = "115年醫師病房值班.xlsx"
df = None

if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='醫師總表')
        df.columns = df.columns.astype(str).str.strip()
        df.fillna({'平日應值班數': 0, '假日應值班數': 0, '固定值班日期': '', 'MVPN': ''}, inplace=True)
    except Exception as e:
        st.error(f"❌ 檔案讀取錯誤：{e}")
        st.stop()
else:
    st.warning(f"⚠️ 請將名為「{file_path}」的檔案上傳至雲端空間。")
    st.stop()

# 建立醫師 MVPN 對應字典
mvpn_dict = {row['醫師姓名']: str(row.get('MVPN', '')).replace('.0', '') for _, row in df.iterrows()}

# 產生日期與星期 (W1~W7) 標籤
days_info = {}
days_display = []
start_date = datetime(2026, 9, 1)
for i in range(30):
    current_date = start_date + timedelta(days=i)
    simple_date = f"{current_date.month}/{current_date.day}" 
    weekday_str = f"W{current_date.weekday() + 1}"
    day_type = "假日" if current_date.weekday() >= 5 else "平日"
    display_str = f"{simple_date} ({weekday_str})"
    days_info[simple_date] = {"類型": day_type, "完整": display_str}
    days_display.append(display_str)

# ==========================================
# 表單送出與自動歸零邏輯 (Callback)
# ==========================================
if 'submit_success' not in st.session_state:
    st.session_state.submit_success = False

def submit_form():
    doc = st.session_state.doc_selector
    days = st.session_state.days_selector
    if doc != "請選擇...":
        # 1. 將資料存入實體 JSON 檔案，確保大家都能看到
        current_prefs = load_preferences()
        current_prefs[doc] = days
        save_preferences(current_prefs)
        
        # 2. 將畫面狀態強制歸零，並觸發成功訊息
        st.session_state.submit_success = True
        st.session_state.doc_selector = "請選擇..."
        st.session_state.days_selector = []

# ==========================================
# 前台：醫師意願登記介面
# ==========================================
if st.session_state.submit_success:
    st.success("✅ 意願已成功送出！畫面已為您重新歸零。")
    st.session_state.submit_success = False

st.subheader("登記值班意願")

first_line_docs = df[df['班別'] == '一線']['醫師姓名'].tolist()
second_line_docs = df[df['班別'] == '二線']['醫師姓名'].tolist()
special_second_line = [doc for doc in ["林中華", "林尚華"] if doc in second_line_docs]
priority_group = first_line_docs + special_second_line

# 計算目前每個日期被選擇的總次數
date_counts = {d: 0 for d in days_display}
for doc, prefs in all_prefs.items():
    for p in prefs:
        if p in date_counts:
            date_counts[p] += 1

selected_doctor = st.selectbox(
    "請選擇您的姓名：", 
    ["請選擇..."] + priority_group, 
    key="doc_selector" # 綁定 Key 供歸零使用
)

if selected_doctor != "請選擇...":
    doc_info = df[df['醫師姓名'] == selected_doctor].iloc[0]
    fixed_dates_str = str(doc_info.get('固定值班日期', ''))
    if fixed_dates_str.strip():
        st.warning(f"📌 提醒：您的固定值班日期為 👉 **{fixed_dates_str}**")
        
    preferred_days = st.multiselect(
        "請勾選您【希望值班】的日期：", 
        options=days_display,
        format_func=lambda x: f"{x} ｜ 目前有 {date_counts.get(x, 0)} 人選擇",
        key="days_selector" # 綁定 Key 供歸零使用
    )
    
    st.button("送出意願", on_click=submit_form)

# ==========================================
# 隱藏後台：密碼解鎖區 (位於左側邊欄)
# ==========================================
with st.sidebar:
    st.write("🔧 **管理員專區**")
    st.write("此區僅供排班人員使用。")
    # 設定解鎖密碼
    admin_password = st.text_input("請輸入密碼解鎖排班後台：", type="password")

    if admin_password == "1234":
        st.success("✅ 後台已解鎖！")
        st.write("---")
        
        if st.button("產生最終班表", type="primary"):
            schedule = {date: None for date in days_info.keys()}
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            
            # 優先級 0：固定值班
            for idx, row in df.iterrows():
                doc_name = row['醫師姓名']
                fixed_str = str(row.get('固定值班日期', ''))
                if fixed_str.strip():
                    dates = [d.strip() for d in fixed_str.split(',')]
                    for d in dates:
                        if d in schedule:
                            schedule[d] = doc_name
            
            # 階段一：一線
            for d, doc in schedule.items():
                if doc is None:
                    interested_first = [doc for doc in first_line_docs if doc in all_prefs and days_info[d]["完整"] in all_prefs[doc]]
                    if interested_first: schedule[d] = random.choice(interested_first)
            
            # 階段二：特定二線
            for d, doc in schedule.items():
                if doc is None:
                    interested_special = [doc for doc in special_second_line if doc in all_prefs and days_info[d]["完整"] in all_prefs[doc]]
                    if interested_special: schedule[d] = random.choice(interested_special)
            
            # 階段三：一般二線
            for d, doc in schedule.items():
                if doc is None:
                    if regular_second_line: 
                        schedule[d] = random.choice(regular_second_line)
                    else: 
                        schedule[d] = "待補"

            # 整理並輸出 Excel
            final_schedule_list = []
            for d, doc in schedule.items():
                display_name = doc
                if doc in mvpn_dict and mvpn_dict[doc] != "":
                    display_name = f"{doc} ({mvpn_dict[doc]})"
                
                final_schedule_list.append({
                    "日期": d,
                    "星期": days_info[d]["完整"].split(" ")[1].replace("(", "").replace(")", ""),
                    "平假日": days_info[d]["類型"],
                    "值班醫師": display_name
                })
                
            df_result = pd.DataFrame(final_schedule_list)
            st.dataframe(df_result, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='最終班表')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載最終班表 (Excel 格式)",
                data=excel_data,
                file_name='醫療部_最終班表.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
