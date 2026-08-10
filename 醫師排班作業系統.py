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
# 網頁基礎與視覺設定 
# ==========================================
st.set_page_config(page_title="醫療部病房排班系統", layout="centered", initial_sidebar_state="collapsed")
st.markdown("""
<style>
    .stApp { background-color: #FAF9F6 !important; }
    [data-testid="stSidebar"] { background-color: #F5F2EA !important; }
    h1, h2, h3, h4, h5, h6, p, div, span, label, li { 
        color: #2C2C2C !important; 
        font-family: "Microsoft JhengHei", sans-serif !important; 
    }
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
    header[data-testid="stHeader"] { display: none !important; }
    [data-testid="collapsedControl"] { display: none !important; }
    [data-testid="stTextInput"] button { display: none !important; }
    button[title="Show password text"] { display: none !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; border-bottom: 2px solid #D3D3D3; }
    .stTabs [data-baseweb="tab"] { height: 50px; background-color: transparent; border: none; }
    .stTabs [aria-selected="true"] { background-color: transparent; border-bottom: 3px solid #333333; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 系統時間與動態月份計算
# ==========================================
today = datetime.today()

if today.month == 12:
    target_year = today.year + 1
    target_month = 1
else:
    target_year = today.year
    target_month = today.month + 1

_, num_days = calendar.monthrange(target_year, target_month)
month_key = f"{target_year}_{target_month}"
config_key = f"config_{month_key}"

# ==========================================
# 核心機制：基礎本機存檔 (JSON)
# ==========================================
PREFS_FILE = 'preferences.json'

def load_all_data():
    if os.path.exists(PREFS_FILE):
        with open(PREFS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

all_json_data = load_all_data()
all_prefs = all_json_data.get(month_key, {})
month_config = all_json_data.get(config_key, {"manual_lock": False})
is_manually_locked = month_config.get("manual_lock", False)

is_open_for_submission = (today.day <= 10) and not is_manually_locked

def save_preferences(prefs):
    data = load_all_data()
    data[month_key] = prefs
    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

def set_manual_lock(status):
    data = load_all_data()
    data[config_key] = {"manual_lock": status}
    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

# ==========================================
# 頂部 Logo 與標題
# ==========================================
logo_path = "yumin_logo.png"
col1, col2 = st.columns([1, 6])
with col1:
    if os.path.exists(logo_path):
        st.image(logo_path, width=80)
with col2:
    st.title("醫療部病房值班排班系統")

# ==========================================
# 公告與提醒事項區塊
# ==========================================
if is_open_for_submission:
    st.info(
        f"📢 **公告：** 一線醫師開放填寫日期為每月 1 日至 10 日，目前開放登記【{target_year} 年 {target_month} 月】之班表。\n\n"
        "💡 **提醒事項：**\n"
        "1. 請於約班時提供 2～3 個可值班時段，以利後續統整。\n"
        "2. 若有其他問題，請洽行政部 琇雯 分機 7115。"
    )
else:
    st.error(
        f"🔒 **公告：** 目前非開放填寫時間（或表單已由管理員鎖定）。【{target_year} 年 {target_month} 月】的表單已關閉。\n\n"
        "💡 **提醒事項：**\n"
        "若有其他問題，請洽行政部 琇雯 分機 7115。"
    )

# ==========================================
# 自動讀取後台檔案
# ==========================================
file_path = "115年醫師病房值班.xlsx"
df = None

if os.path.exists(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='醫師總表')
        df.columns = df.columns.astype(str).str.strip()
        required_cols = ['醫師姓名', '班別', '平日應值班數', '假日應值班數', '排班規則', 'MVPN']
        for col in required_cols:
            if col not in df.columns:
                df[col] = ''
        
        df['平日應值班數'] = pd.to_numeric(df['平日應值班數'], errors='coerce').fillna(0).astype(int)
        df['假日應值班數'] = pd.to_numeric(df['假日應值班數'], errors='coerce').fillna(0).astype(int)
        df.fillna('', inplace=True)
    except Exception as e:
        st.error(f"❌ 檔案讀取錯誤：{e}")
        st.stop()
else:
    st.error(f"⚠️ [系統管理員請注意]：後台找不到名為「{file_path}」的 Excel 檔案。請確認檔案是否存在。")
    st.stop()

# 建立字典供後續快速查詢
all_doctors = df['醫師姓名'].dropna().unique().tolist()
mvpn_dict = {row['醫師姓名']: str(row['MVPN']).replace('.0', '') for _, row in df.iterrows()}
line_type_dict = {row['醫師姓名']: str(row['班別']).strip() for _, row in df.iterrows()}

first_line_docs = df[df['班別'] == '一線']['醫師姓名'].tolist()
second_line_docs = df[df['班別'] == '二線']['醫師姓名'].tolist()

# 單雙月規則過濾
is_odd_month = (target_month % 2 != 0)
first_line_docs = [doc for doc in first_line_docs if not (is_odd_month and doc == "陳儀聲") and not (not is_odd_month and doc == "詹鈞惟")]
second_line_docs = [doc for doc in second_line_docs if not (is_odd_month and doc == "陳儀聲") and not (not is_odd_month and doc == "詹鈞惟")]
priority_group = first_line_docs + [doc for doc in ["林中華", "林尚華"] if doc in second_line_docs]

# ==========================================
# 日期與星期計算
# ==========================================
days_info = {}
days_display = []
start_date = datetime(target_year, target_month, 1)
weekday_zh_map = {1: "星期一", 2: "星期二", 3: "星期三", 4: "星期四", 5: "星期五", 6: "星期六", 7: "星期日"}

for i in range(num_days):
    current_date = start_date + timedelta(days=i)
    simple_date = f"{current_date.month}/{current_date.day}" 
    weekday_num = current_date.weekday() + 1
    weekday_str = f"W{weekday_num}"
    zh_weekday = weekday_zh_map[weekday_num]
    day_type = "假日" if weekday_num >= 6 else "平日"
    display_str = f"{simple_date} ({weekday_str})"
    
    days_info[simple_date] = {"類型": day_type, "完整": display_str, "星期": weekday_str, "中文星期": zh_weekday}
    days_display.append(display_str)

date_counts = {d: 0 for d in days_display}
for doc, prefs in all_prefs.items():
    for p in prefs:
        if p in date_counts:
            date_counts[p] += 1

# ==========================================
# 前台與後台：三頁籤設計
# ==========================================
if 'unlock_edit' not in st.session_state:
    st.session_state.unlock_edit = False

def reset_unlock_state():
    st.session_state.unlock_edit = False

tab1, tab2, tab3 = st.tabs(["📝 登記與修改意願", "📊 查詢未登記名單", "⚙️ 管理員排班後台"])

# ------------------------------------------
# 頁籤 1: 登記與修改
# ------------------------------------------
with tab1:
    st.subheader(f"登記或修改 {target_month} 月份值班意願")
    selected_doctor = st.selectbox("請選擇您的姓名：", ["請選擇..."] + priority_group, on_change=reset_unlock_state)

    if selected_doctor != "請選擇...":
        existing_prefs = all_prefs.get(selected_doctor, [])
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
                st.info("🔒 您的意願已儲存並鎖定。若需修改，請點擊下方按鈕解鎖。")
                if st.button("修改意願", type="primary"):
                    st.session_state.unlock_edit = True
                    st.rerun()
            else:
                if st.button("儲存意願", type="primary"):
                    all_prefs[selected_doctor] = preferred_days
                    save_preferences(all_prefs)
                    st.session_state.unlock_edit = False
                    st.success("✅ 意願已成功儲存！系統已自動鎖定您的表單。")
                    time.sleep(1)
                    st.rerun()

# ------------------------------------------
# 頁籤 2: 查詢未登記
# ------------------------------------------
with tab2:
    st.subheader(f"尚未登記 {target_month} 月份意願的醫師名單")
    unsubmitted_docs = [doc for doc in priority_group if doc not in all_prefs]
    if unsubmitted_docs:
        st.warning(f"目前共有 **{len(unsubmitted_docs)}** 位醫師尚未登記：")
        for doc in unsubmitted_docs:
            st.write(f"- {doc}")
    else:
        st.success("🎉 太棒了！所有優先群組的醫師皆已完成意願登記。")

# ------------------------------------------
# 頁籤 3: 管理員排班後台
# ------------------------------------------
with tab3:
    st.write("🔧 **解鎖排班管理系統**")
    admin_password = st.text_input("請輸入管理員密碼：", type="password")

    if admin_password == "914192":
        st.success("✅ 後台已解鎖！")
        st.markdown("---")
        
        st.subheader("📋 醫師意願清單總覽")
        prefs_summary_data = []
        for doc in priority_group:
            choices = all_prefs.get(doc, [])
            prefs_summary_data.append({
                "醫師姓名": doc,
                "登記意願 (日期)": ", ".join(choices) if choices else "尚未登記或無意願"
            })
        df_prefs = pd.DataFrame(prefs_summary_data)
        st.dataframe(df_prefs, use_container_width=True)
        st.markdown("---")
        
        st.subheader("✏️ 手動調整醫師意願")
        st.write("無視前台鎖定狀態，管理員可在此直接修改醫師的約班日期。")
        admin_selected_doc = st.selectbox("選擇要調整的醫師：", ["請選擇..."] + priority_group, key="admin_doc")
        
        if admin_selected_doc != "請選擇...":
            admin_current_prefs = all_prefs.get(admin_selected_doc, [])
            admin_new_prefs = st.multiselect(
                "調整該醫師的意願日期：", 
                options=days_display, 
                default=admin_current_prefs, 
                key="admin_multi"
            )
            
            if st.button(f"儲存 {admin_selected_doc} 的調整", type="primary"):
                all_prefs[admin_selected_doc] = admin_new_prefs
                save_preferences(all_prefs)
                st.success(f"✅ 已成功強制更新 {admin_selected_doc} 的意願！")
                time.sleep(1)
                st.rerun()

        st.markdown("---")
        
        st.subheader("🔒 系統狀態與危險操作")
        colA, colB = st.columns(2)
        
        with colA:
            st.write("**【表單鎖定控制】**")
            if is_manually_locked:
                st.warning("目前前台已被 **手動強制鎖定**。")
                if st.button("🔓 解鎖前台表單"):
                    set_manual_lock(False)
                    st.rerun()
            else:
                st.info("目前前台依據日期機制正常運作中。")
                if st.button("🔐 強制鎖定前台表單"):
                    set_manual_lock(True)
                    st.rerun()
                    
        with colB:
            st.write("**【資料重設】**")
            if st.button("🗑️ 重設當月所有意願紀錄"):
                data = load_all_data()
                if month_key in data and len(data[month_key]) > 0:
                    data[month_key] = {}
                    with open(PREFS_FILE, 'w', encoding='utf-8') as f:
                        json.dump(data, f, ensure_ascii=False)
                    st.success("✅ 當月意願紀錄已全數歸零！")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.info("目前尚無登記紀錄，無需歸零。")
        
        st.markdown("---")
        
        st.subheader("🗓️ 智慧排班演算法")
        
        if st.button("執行排班並產生 Excel", type="primary"):
            # 每天只有 1 個位子
            schedule = {date: None for date in days_info.keys()}
            
            # 讀取每位醫師的應值班數限制
            rem_quota = {doc: {"平日": int(df.loc[df['醫師姓名']==doc, '平日應值班數'].values[0]),
                               "假日": int(df.loc[df['醫師姓名']==doc, '假日應值班數'].values[0])} 
                         for doc in all_doctors if doc in df['醫師姓名'].values}
            assigned_counts = {doc: {"平日": 0, "假日": 0} for doc in all_doctors}
            
            special_second_line = [doc for doc in ["林中華", "林尚華"] if doc in second_line_docs]
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            priority_docs = first_line_docs + special_second_line

            # 【階段一】嚴格依據意願分配 (優先名單：一線 + 特殊二線)
            made_progress = True
            while made_progress:
                made_progress = False
                doc_avail_choices = {}
                
                for doc in priority_docs:
                    if rem_quota[doc]["平日"] <= 0 and rem_quota[doc]["假日"] <= 0:
                        continue 
                    
                    prefs_display = all_prefs.get(doc, [])
                    valid_dates = []
                    for pref_str in prefs_display:
                        for d, info in days_info.items():
                            if info["完整"] == pref_str:
                                if schedule[d] is None and rem_quota[doc][info["類型"]] > 0:
                                    valid_dates.append(d)
                                break
                    if valid_dates:
                        doc_avail_choices[doc] = valid_dates
                
                if doc_avail_choices:
                    sorted_docs = sorted(doc_avail_choices.keys(), key=lambda x: len(doc_avail_choices[x]))
                    target_doc = sorted_docs[0]
                    
                    day_contention = {}
                    for d in doc_avail_choices[target_doc]:
                        contention = sum(1 for other_doc, choices in doc_avail_choices.items() if d in choices)
                        day_contention[d] = contention
                        
                    best_day = sorted(doc_avail_choices[target_doc], key=lambda x: day_contention[x])[0]
                    
                    schedule[best_day] = target_doc
                    day_type = days_info[best_day]["類型"]
                    rem_quota[target_doc][day_type] -= 1
                    assigned_counts[target_doc][day_type] += 1
                    made_progress = True

            # 【階段二】剩餘空白日期，分配給一般二線醫師 (週間與假日獨立演算 + 平假日互斥)
            empty_days = [d for d, doc in schedule.items() if doc is None]
            
            for d in empty_days:
                day_type = days_info[d]["類型"] 
                zh_weekday = days_info[d]["中文星期"]
                
                available_second = []
                for doc in regular_second_line:
                    rules = str(df.loc[df['醫師姓名']==doc, '排班規則'].values[0])
                    if doc == "李友夫" and zh_weekday in ["星期二", "星期三", "星期四", "星期五"] and "排除星期二到星期五" in rules:
                        continue
                    
                    # 💡 核心更新：平假日互斥邏輯
                    # 如果當前空缺是平日，且該醫師已經有假日班，則跳過
                    if day_type == "平日" and assigned_counts[doc]["假日"] > 0:
                        continue
                    # 如果當前空缺是假日，且該醫師已經有平日班，則跳過
                    if day_type == "假日" and assigned_counts[doc]["平日"] > 0:
                        continue

                    available_second.append(doc)
                
                # 💡 防呆機制：如果因為互斥條件導致這天完全沒人可選，則放寬條件以防系統出錯
                if not available_second:
                    for doc in regular_second_line:
                        rules = str(df.loc[df['醫師姓名']==doc, '排班規則'].values[0])
                        if doc == "李友夫" and zh_weekday in ["星期二", "星期三", "星期四", "星期五"] and "排除星期二到星期五" in rules:
                            continue
                        available_second.append(doc)

                if available_second:
                    min_shifts = min([assigned_counts[doc][day_type] for doc in available_second])
                    candidates = [doc for doc in available_second if assigned_counts[doc][day_type] == min_shifts]
                    chosen = random.choice(candidates)
                    schedule[d] = chosen
                    assigned_counts[chosen][day_type] += 1

            # 整理輸出格式
            final_schedule_list = []
            for d, doc in schedule.items():
                doc1, doc2, mvpn1, mvpn2 = "", "", "", ""
                
                if doc:
                    doc_type = line_type_dict.get(doc, "")
                    if doc_type == "一線":
                        doc1 = doc
                        mvpn1 = f"{doc}{mvpn_dict.get(doc, '')}"
                    elif doc_type == "二線":
                        doc2 = doc
                        mvpn2 = f"{doc}{mvpn_dict.get(doc, '')}"
                
                final_schedule_list.append({
                    "日期": d,
                    "星期": days_info[d]["星期"],
                    "一線": doc1,
                    "二線": doc2,
                    "一線MVPN": mvpn1,
                    "二線MVPN": mvpn2
                })
                
            df_result = pd.DataFrame(final_schedule_list)
            st.success("✨ 班表運算完成！完全依照醫師預約意願與配額進行排班。")
            st.dataframe(df_result, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name=f'{target_month}月最終班表')
                df_prefs.to_excel(writer, index=False, sheet_name='醫師意願清單')
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 下載完整資料 (含班表與意願清單)",
                data=excel_data,
                file_name=f'醫療部_{target_year}年{target_month}月_排班總表.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
