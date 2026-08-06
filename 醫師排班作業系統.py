import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
import io
import os
import json
import time

# ==========================================
# 網頁基礎與視覺設定 (暖奶油色背景、文字為主)
# ==========================================
st.set_page_config(page_title="醫療部病房排班系統", layout="centered")
st.markdown("""
<style>
    .stApp { background-color: #FDFBF7; }
    h1, h2, h3, p, div, span, label { color: #333333; font-family: sans-serif; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px 4px 0px 0px; gap: 1px; padding-top: 10px; padding-bottom: 10px; }
    .stTabs [aria-selected="true"] { background-color: #EFE9D9; border-bottom: 2px solid #555555; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

st.title("醫療部病房值班排班系統")
st.info("📢 **公告：一線醫師開放填寫日期為每月 1 日至 10 日，請於期限內完成登記。**")

# ==========================================
# 核心機制：多人共用資料庫 (JSON 即時存檔與覆蓋)
# ==========================================
PREFS_FILE = 'preferences.json'

def load_preferences():
    # 讀取目前存檔的資料
    if os.path.exists(PREFS_FILE):
        with open(PREFS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_preferences(prefs):
    # 將最新資料寫入檔案保存
    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
        json.dump(prefs, f, ensure_ascii=False)

# 每次網頁刷新時，都會抓取最新存檔的資料
all_prefs = load_preferences()

# ==========================================
# 自動讀取雲端後台檔案
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
    st.error(f"⚠️ [系統管理員請注意]：後台找不到名為「{file_path}」的 Excel 檔案。請上傳至雲端空間。")
    st.stop()

# 建立醫師基礎資料字典
all_doctors = df['醫師姓名'].dropna().unique().tolist()
mvpn_dict = {row['醫師姓名']: str(row.get('MVPN', '')).replace('.0', '') for _, row in df.iterrows()}
quota_dict = {row['醫師姓名']: {'平日': int(row['平日應值班數']), '假日': int(row['假日應值班數'])} for _, row in df.iterrows()}
line_type_dict = {row['醫師姓名']: str(row['班別']).strip() for _, row in df.iterrows()}

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

# 計算目前每個日期被選擇的總次數
date_counts = {d: 0 for d in days_display}
for doc, prefs in all_prefs.items():
    for p in prefs:
        if p in date_counts:
            date_counts[p] += 1

# ==========================================
# 前台介面：雙頁籤設計
# ==========================================
tab1, tab2 = st.tabs(["📝 登記與修改意願", "📊 查詢未登記名單"])

with tab1:
    st.subheader("登記或修改您的值班意願")
    
    first_line_docs = df[df['班別'] == '一線']['醫師姓名'].tolist()
    second_line_docs = df[df['班別'] == '二線']['醫師姓名'].tolist()
    special_second_line = [doc for doc in ["林中華", "林尚華"] if doc in second_line_docs]
    priority_group = first_line_docs + special_second_line

    selected_doctor = st.selectbox("請選擇您的姓名：", ["請選擇..."] + priority_group)

    if selected_doctor != "請選擇...":
        doc_info = df[df['醫師姓名'] == selected_doctor].iloc[0]
        fixed_dates_str = str(doc_info.get('固定值班日期', ''))
        if fixed_dates_str.strip():
            st.warning(f"📌 提醒：您的固定值班日期為 👉 **{fixed_dates_str}**")
            
        # 讀取該醫師目前的意願，作為選單的預設值 (顯示最新狀態)
        existing_prefs = all_prefs.get(selected_doctor, [])
        
        preferred_days = st.multiselect(
            "請勾選或修改您【希望值班】的日期：", 
            options=days_display,
            format_func=lambda x: f"{x} ｜ 目前有 {date_counts.get(x, 0)} 人選擇",
            default=existing_prefs
        )
        
        if st.button("儲存意願", type="primary"):
            # 讀取最新資料 -> 覆蓋該醫師的選擇 -> 重新存檔
            current_prefs = load_preferences()
            current_prefs[selected_doctor] = preferred_days 
            save_preferences(current_prefs)
            
            st.success("✅ 意願已成功儲存！系統已紀錄您最新的選擇。")
            time.sleep(1) # 暫停 1 秒讓醫師看到成功訊息
            st.rerun()    # 重新整理畫面，確保資料同步

with tab2:
    st.subheader("尚未登記意願的醫師名單")
    unsubmitted_docs = [doc for doc in priority_group if doc not in all_prefs or not all_prefs[doc]]
    
    if unsubmitted_docs:
        st.warning(f"目前共有 **{len(unsubmitted_docs)}** 位醫師尚未登記：")
        for doc in unsubmitted_docs:
            st.write(f"- {doc}")
    else:
        st.success("🎉 太棒了！所有優先群組的醫師皆已完成意願登記。")

# ==========================================
# 隱藏後台：側邊欄密碼解鎖區
# ==========================================
with st.sidebar:
    st.write("🔧 **管理員專區**")
    admin_password = st.text_input("請輸入密碼解鎖排班後台：", type="password")

    if admin_password == "1234":
        st.success("✅ 後台已解鎖！")
        
        if st.button("產生最終班表", type="primary"):
            # 每天只有【一位】醫師值班的邏輯
            schedule = {date: None for date in days_info.keys()}
            assigned_counts = {doc: {"平日": 0, "假日": 0} for doc in all_doctors}
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            
            # 【優先級 0】固定值班日期
            for _, row in df.iterrows():
                doc_name = row['醫師姓名']
                fixed_str = str(row.get('固定值班日期', ''))
                if fixed_str.strip():
                    for d in [x.strip() for x in fixed_str.split(',')]:
                        if d in schedule and schedule[d] is None:
                            schedule[d] = doc_name
                            assigned_counts[doc_name][days_info[d]["類型"]] += 1
            
            # 【階段一】智慧排班：一線醫師
            for d in schedule.keys():
                if schedule[d] is None:
                    day_type = days_info[d]["類型"]
                    full_day_str = days_info[d]["完整"]
                    interested = [doc for doc in first_line_docs if doc in all_prefs and full_day_str in all_prefs[doc]]
                    valid_candidates = [doc for doc in interested if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d] = chosen
                        assigned_counts[chosen][day_type] += 1
                    elif interested:
                        chosen = random.choice(interested)
                        schedule[d] = chosen
                        assigned_counts[chosen][day_type] += 1

            # 【階段二】特定二線優先
            for d in schedule.keys():
                if schedule[d] is None:
                    day_type = days_info[d]["類型"]
                    full_day_str = days_info[d]["完整"]
                    interested = [doc for doc in special_second_line if doc in all_prefs and full_day_str in all_prefs[doc]]
                    valid_candidates = [doc for doc in interested if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d] = chosen
                        assigned_counts[chosen][day_type] += 1

            # 【階段三】一般二線填補空缺
            for d in schedule.keys():
                if schedule[d] is None:
                    day_type = days_info[d]["類型"]
                    valid_candidates = [doc for doc in regular_second_line if assigned_counts[doc][day_type] < quota_dict[doc][day_type]]
                    if valid_candidates:
                        chosen = random.choice(valid_candidates)
                        schedule[d] = chosen
                        assigned_counts[chosen][day_type] += 1
                    else:
                        schedule[d] = random.choice(regular_second_line) if regular_second_line else ""

            # 整理輸出，確保一線與二線不會同時出現在同一天
            final_schedule_list = []
            for d, doc in schedule.items():
                doc_type = line_type_dict.get(doc, "")
                
                doc1 = doc if doc_type == "一線" else ""
                doc2 = doc if doc_type == "二線" else ""
                
                mvpn1 = f"{doc1}{mvpn_dict.get(doc1, '')}" if doc1 else ""
                mvpn2 = f"{doc2}{mvpn_dict.get(doc2, '')}" if doc2 else ""
                
                if doc and not doc1 and not doc2:
                    doc1, mvpn1 = doc, f"{doc}{mvpn_dict.get(doc, '')}"
                
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
