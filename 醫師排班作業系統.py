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

all_prefs = load_preferences()

# ==========================================
# 自動讀取雲端檔案 (115年醫師病房值班.xlsx)
# ==========================================
file_path = "115年醫師病房值班.xlsx"
df = None

if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='醫師總表')
        df.columns = df.columns.astype(str).str.strip()
        df.fillna({'平日應值班數': 0, '假日應值班數': 0, '固定值班日期': '', 'MVPN': ''}, inplace=True)
    except Exception as e:
        st.error(f"❌ 檔案讀取錯誤，請確認檔案是否有「醫師總表」分頁：{e}")
        st.stop()
else:
    st.warning(f"⚠️ 系統準備就緒。請確認雲端空間內有上傳「{file_path}」。")
    st.stop()

# 建立 MVPN 與配額字典
mvpn_dict = {row['醫師姓名']: str(row.get('MVPN', '')).replace('.0', '') for _, row in df.iterrows()}
quota_dict = {row['醫師姓名']: {'平日': int(row['平日應值班數']), '假日': int(row['假日應值班數'])} for _, row in df.iterrows()}

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
    
    days_info[simple_date] = {"類型": day_type, "完整": display_str, "星期": weekday_str}
    days_display.append(display_str)

# ==========================================
# 表單送出與自動歸零邏輯 (Callback)
# ==========================================
if 'submit_success' not in st.session_state:
    st.session_state.submit_success = False

# 初始化選單的 session_state 以防報錯
if 'doc_selector' not in st.session_state:
    st.session_state.doc_selector = "請選擇..."
if 'days_selector' not in st.session_state:
    st.session_state.days_selector = []

def submit_form():
    doc = st.session_state.doc_selector
    days = st.session_state.days_selector
    if doc != "請選擇...":
        current_prefs = load_preferences()
        current_prefs[doc] = days
        save_preferences(current_prefs)
        
        st.session_state.submit_success = True
        # 強制歸零
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

date_counts = {d: 0 for d in days_display}
for doc, prefs in all_prefs.items():
    for p in prefs:
        if p in date_counts:
            date_counts[p] += 1

st.selectbox("請選擇您的姓名：", ["請選擇..."] + priority_group, key="doc_selector")

if st.session_state.doc_selector != "請選擇...":
    doc_info = df[df['醫師姓名'] == st.session_state.doc_selector].iloc[0]
    fixed_dates_str = str(doc_info.get('固定值班日期', ''))
    if fixed_dates_str.strip():
        st.warning(f"📌 提醒：您的固定值班日期為 👉 **{fixed_dates_str}**")
        
    st.multiselect(
        "請勾選您【希望值班】的日期：", 
        options=days_display,
        format_func=lambda x: f"{x} ｜ 目前有 {date_counts.get(x, 0)} 人選擇",
        key="days_selector"
    )
    
    st.button("送出意願", on_click=submit_form)

# ==========================================
# 隱藏後台：側邊欄密碼解鎖區
# ==========================================
with st.sidebar:
    st.write("🔧 **管理員專區**")
    admin_password = st.text_input("請輸入密碼解鎖排班後台：", type="password")

    if admin_password == "1234":
        st.success("✅ 後台已解鎖！")
        
        if st.button("產生最終班表", type="primary"):
            # 初始化排班表：每天區分一線與二線
            schedule = {date: {"一線": None, "二線": None} for date in days_info.keys()}
            assigned_counts = {doc: {"平日": 0, "假日": 0} for doc in df['醫師姓名']}
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            
            # 【優先級 0】固定值班日期
            for _, row in df.iterrows():
                doc_name = row['醫師姓名']
                line_type = row['班別']
                fixed_str = str(row.get('固定值班日期', ''))
                if fixed_str.strip():
                    for d in [x.strip() for x in fixed_str.split(',')]:
                        if d in schedule:
                            schedule[d][line_type] = doc_name
                            assigned_counts[doc_name][days_info[d]["類型"]] += 1
            
            # 【階段一】智慧排班：一線醫師
            for d, info in schedule.items():
                if info["一線"] is None:
                    day_type = days_info[d]["類型"]
                    full_day_str = days_info[d]["完整"]
                    # 找出有意願且未超過該日類型配額的一線醫師
                    interested = [doc for doc in first_line_docs if doc in all_prefs and full_day_str in all_prefs[doc]]
                    valid_candidates = [doc for doc in interested if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d]["一線"] = chosen
                        assigned_counts[chosen][day_type] += 1
                    elif interested:
                        # 如果大家都滿了但還是有人選，強制隨機選一個
                        chosen = random.choice(interested)
                        schedule[d]["一線"] = chosen
                        assigned_counts[chosen][day_type] += 1

            # 【階段二】特定二線優先
            for d, info in schedule.items():
                if info["二線"] is None:
                    day_type = days_info[d]["類型"]
                    full_day_str = days_info[d]["完整"]
                    interested = [doc for doc in special_second_line if doc in all_prefs and full_day_str in all_prefs[doc]]
                    valid_candidates = [doc for doc in interested if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d]["二線"] = chosen
                        assigned_counts[chosen][day_type] += 1

            # 【階段三】一般二線填補空缺
            for d, info in schedule.items():
                if info["二線"] is None:
                    day_type = days_info[d]["類型"]
                    valid_candidates = [doc for doc in regular_second_line if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d]["二線"] = chosen
                        assigned_counts[chosen][day_type] += 1
                    else:
                        schedule[d]["二線"] = random.choice(regular_second_line) if regular_second_line else ""

            # 整理輸出符合班表呈現的格式
            final_schedule_list = []
            for d, info in schedule.items():
                doc1 = info["一線"] if info["一線"] else ""
                doc2 = info["二線"] if info["二線"] else ""
                
                # 組合 MVPN 字串 (如：李毓珊678910)
                mvpn1 = f"{doc1}{mvpn_dict.get(doc1, '')}" if doc1 else ""
                mvpn2 = f"{doc2}{mvpn_dict.get(doc2, '')}" if doc2 else ""
                
                final_schedule_list.append({
                    "日期": d,
                    "星期": days_info[d]["星期"],
                    "一線": doc1,
                    "二線": doc2,
                    "一線MVPN": mvpn1,
                    "二線MVPN": mvpn2
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
                file_name='115年醫師病房值班_排班結果.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
