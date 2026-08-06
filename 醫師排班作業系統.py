import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
import calendar
import io
import os
import json
import time

# ==========================================
# 網頁基礎與視覺設定 (溫暖奶油白、馬卡龍綠、隱藏殘影)
# ==========================================
st.set_page_config(page_title="醫療部病房排班系統", layout="centered", initial_sidebar_state="expanded")
st.markdown("""
<style>
    .stApp { background-color: #FAF9F6 !important; }
    [data-testid="stSidebar"] { background-color: #F5F2EA !important; }
    h1, h2, h3, h4, h5, h6, p, div, span, label, li { 
        color: #2C2C2C !important; 
        font-family: "Microsoft JhengHei", sans-serif !important; 
    }
    
    /* 多選選單標籤精準覆蓋 (馬卡龍淺綠色) */
    div[data-baseweb="select"] span[data-baseweb="tag"] {
        background-color: #D4EFDF !important; 
        color: #333333 !important;
        border: none !important;
    }
    div[data-baseweb="select"] span[data-baseweb="tag"] * {
        background-color: transparent !important;
        color: #333333 !important;
    }
    div[data-baseweb="select"] span[data-baseweb="tag"] svg {
        fill: #333333 !important;
    }
    div[data-baseweb="select"] > div:focus-within {
        border-color: #D4EFDF !important;
        box-shadow: 0 0 0 2px #D4EFDF !important;
    }
    
    /* 主要按鈕 (燕麥/淺沙色) */
    button[kind="primary"] {
        background-color: #E2D9C8 !important;
        border: 1px solid #D1C7B4 !important;
        color: #333333 !important;
    }
    button[kind="primary"]:hover {
        background-color: #D1C7B4 !important;
        border: 1px solid #C0B5A1 !important;
        color: #333333 !important;
    }
    
    /* 隱藏頂部控制列與圖示殘影 */
    header[data-testid="stHeader"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    [data-testid="stSidebarCollapseButton"] { display: none !important; }
    [data-testid="stTextInput"] button { display: none !important; }
    button[title="Show password text"] { display: none !important; }
    
    /* 頁籤設計 */
    .stTabs [data-baseweb="tab-list"] { gap: 24px; border-bottom: 2px solid #D3D3D3; }
    .stTabs [data-baseweb="tab"] { height: 50px; background-color: transparent; border: none; }
    .stTabs [aria-selected="true"] { background-color: transparent; border-bottom: 3px solid #333333; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 系統時間與動態月份計算
# ==========================================
today = datetime.today()
is_open_for_submission = today.day <= 10

if today.month == 12:
    target_year = today.year + 1
    target_month = 1
else:
    target_year = today.year
    target_month = today.month + 1

_, num_days = calendar.monthrange(target_year, target_month)
month_key = f"{target_year}_{target_month}"

# ==========================================
# 頂部 Logo 與標題顯示區塊
# ==========================================
logo_path = "yumin_logo.png"
col1, col2 = st.columns([1, 6])
with col1:
    if os.path.exists(logo_path):
        st.image(logo_path, width=80)
with col2:
    st.title("醫療部病房值班排班系統")

if is_open_for_submission:
    st.info(f"📢 **公告：一線醫師開放填寫日期為每月 1 日至 10 日，目前開放登記【{target_year} 年 {target_month} 月】之班表。**")
else:
    st.error(f"🔒 **公告：目前非開放填寫時間。【{target_year} 年 {target_month} 月】的表單已關閉。**")

# ==========================================
# 核心機制：多人共用資料庫 (依月份獨立存檔)
# ==========================================
PREFS_FILE = 'preferences.json'

def load_preferences():
    if os.path.exists(PREFS_FILE):
        with open(PREFS_FILE, 'r', encoding='utf-8') as f:
            all_data = json.load(f)
            return all_data.get(month_key, {})
    return {}

def save_preferences(prefs):
    all_data = {}
    if os.path.exists(PREFS_FILE):
        with open(PREFS_FILE, 'r', encoding='utf-8') as f:
            all_data = json.load(f)
    all_data[month_key] = prefs
    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, ensure_ascii=False)

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

# ==========================================
# 名單過濾規則：單雙月顯示限制
# ==========================================
all_doctors = df['醫師姓名'].dropna().unique().tolist()
mvpn_dict = {row['醫師姓名']: str(row.get('MVPN', '')).replace('.0', '') for _, row in df.iterrows()}
quota_dict = {row['醫師姓名']: {'平日': int(row['平日應值班數']), '假日': int(row['假日應值班數'])} for _, row in df.iterrows()}
line_type_dict = {row['醫師姓名']: str(row['班別']).strip() for _, row in df.iterrows()}

first_line_docs = df[df['班別'] == '一線']['醫師姓名'].tolist()
second_line_docs = df[df['班別'] == '二線']['醫師姓名'].tolist()

# 根據目標月份動態過濾
is_odd_month = (target_month % 2 != 0)

# 單數月排除陳儀聲，雙數月排除詹鈞惟
first_line_docs = [doc for doc in first_line_docs if not (is_odd_month and doc == "陳儀聲") and not (not is_odd_month and doc == "詹鈞惟")]
second_line_docs = [doc for doc in second_line_docs if not (is_odd_month and doc == "陳儀聲") and not (not is_odd_month and doc == "詹鈞惟")]

special_second_line = [doc for doc in ["林中華", "林尚華"] if doc in second_line_docs]
priority_group = first_line_docs + special_second_line

# ==========================================
# 日期與星期計算
# ==========================================
days_info = {}
days_display = []
start_date = datetime(target_year, target_month, 1)

for i in range(num_days):
    current_date = start_date + timedelta(days=i)
    simple_date = f"{current_date.month}/{current_date.day}" 
    weekday_str = f"W{current_date.weekday() + 1}"
    day_type = "假日" if current_date.weekday() >= 5 else "平日"
    display_str = f"{simple_date} ({weekday_str})"
    
    days_info[simple_date] = {"類型": day_type, "完整": display_str, "星期": weekday_str}
    days_display.append(display_str)

date_counts = {d: 0 for d in days_display}
for doc, prefs in all_prefs.items():
    for p in prefs:
        if p in date_counts:
            date_counts[p] += 1

# ==========================================
# 前台介面：狀態重置與雙頁籤
# ==========================================
# 初始化解鎖狀態。當姓名選單切換時，呼叫此函數重置解鎖狀態。
if 'unlock_edit' not in st.session_state:
    st.session_state.unlock_edit = False

def reset_unlock_state():
    st.session_state.unlock_edit = False

tab1, tab2 = st.tabs(["📝 登記與修改意願", "📊 查詢未登記名單"])

with tab1:
    st.subheader(f"登記或修改 {target_month} 月份值班意願")

    selected_doctor = st.selectbox(
        "請選擇您的姓名：", 
        ["請選擇..."] + priority_group,
        on_change=reset_unlock_state
    )

    if selected_doctor != "請選擇...":
        doc_info = df[df['醫師姓名'] == selected_doctor].iloc[0]
        fixed_dates_str = str(doc_info.get('固定值班日期', ''))
        if fixed_dates_str.strip():
            st.warning(f"📌 提醒：您的固定值班日期為 👉 **{fixed_dates_str}**")
            
        existing_prefs = all_prefs.get(selected_doctor, [])
        
        # 判斷是否上鎖：該醫師有存檔紀錄，且尚未點擊「解鎖」按鈕
        is_locked = (selected_doctor in all_prefs) and (not st.session_state.unlock_edit)
        
        preferred_days = st.multiselect(
            "請勾選或修改您【希望值班】的日期：", 
            options=days_display,
            format_func=lambda x: f"{x} ｜ 目前有 {date_counts.get(x, 0)} 人選擇",
            default=existing_prefs,
            disabled=is_locked or not is_open_for_submission
        )
        
        if is_open_for_submission:
            if is_locked:
                # 鎖定狀態下，顯示修改按鈕
                st.info("🔒 您的意願已儲存並鎖定。若需修改，請點擊下方按鈕解鎖。")
                if st.button("修改意願", type="primary"):
                    st.session_state.unlock_edit = True
                    st.rerun()
            else:
                # 解鎖狀態下，顯示儲存按鈕
                if st.button("儲存意願", type="primary"):
                    current_prefs = load_preferences()
                    current_prefs[selected_doctor] = preferred_days
                    save_preferences(current_prefs)
                    
                    # 儲存後立即上鎖
                    st.session_state.unlock_edit = False
                    st.success("✅ 意願已成功儲存！系統已自動鎖定您的表單。")
                    time.sleep(1)
                    st.rerun()

with tab2:
    st.subheader(f"尚未登記 {target_month} 月份意願的醫師名單")
    unsubmitted_docs = [doc for doc in priority_group if doc not in all_prefs]
    
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
            schedule = {date: None for date in days_info.keys()}
            assigned_counts = {doc: {"平日": 0, "假日": 0} for doc in all_doctors}
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            
            for _, row in df.iterrows():
                doc_name = row['醫師姓名']
                
                # 確保固定值班的醫師符合單雙月規則才排入
                if (is_odd_month and doc_name == "陳儀聲") or (not is_odd_month and doc_name == "詹鈞惟"):
                    continue
                    
                fixed_str = str(row.get('固定值班日期', ''))
                if fixed_str.strip():
                    for d in [x.strip() for x in fixed_str.split(',')]:
                        if d in schedule and schedule[d] is None:
                            schedule[d] = doc_name
                            assigned_counts[doc_name][days_info[d]["類型"]] += 1
            
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
                df_result.to_excel(writer, index=False, sheet_name=f'{target_month}月最終班表')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載最終班表 (Excel 格式)",
                data=excel_data,
                file_name=f'醫療部_{target_year}年{target_month}月_排班結果.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
