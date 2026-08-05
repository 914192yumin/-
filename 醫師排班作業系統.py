import streamlit as st
import pandas as pd
import random
from datetime import datetime, timedelta
import io

# ==========================================
# 網頁基礎與視覺設定 (暖奶油色背景、乾淨文字排版)
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
# 資料儲存與初始化
# ==========================================
if 'preferences' not in st.session_state:
    st.session_state.preferences = {}
if 'doc_database' not in st.session_state:
    st.session_state.doc_database = None

# ==========================================
# 步驟一：讀取「醫師總表」
# ==========================================
# 這裡已經幫你把顯示文字改成最新的檔名了！
uploaded_file = st.file_uploader("請上傳醫師總表 (115年醫師病房值班.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 指定讀取名為「醫師總表」的工作表
        df = pd.read_excel(uploaded_file, sheet_name='醫師總表')
        df.fillna({'平日應值班數': 0, '假日應值班數': 0, '固定值班日期': '', 'MVPN': ''}, inplace=True)
        st.session_state.doc_database = df
        st.success("✅ 「醫師總表」讀取成功！請透過下方頁籤進行操作。")
    except Exception as e:
        st.error(f"讀取失敗：請確認您的 Excel 檔案中，最下方的工作表(Sheet)名稱是否為「醫師總表」。詳細錯誤：{e}")

# ==========================================
# 建立前台與後台頁籤
# ==========================================
if st.session_state.doc_database is not None:
    df = st.session_state.doc_database
    
    # 建立醫師 MVPN 對應字典
    mvpn_dict = {}
    for _, row in df.iterrows():
        mvpn_val = str(row.get('MVPN', '')).replace('.0', '')
        mvpn_dict[row['醫師姓名']] = mvpn_val

    tab1, tab2 = st.tabs(["👤 前台：醫師意願登記", "⚙️ 後台：自動排班處理"])
    
    # 產生下個月日期與平假日標籤 (以 9 月為例)
    days_info = {}
    days_display = []
    start_date = datetime(2026, 9, 1)
    for i in range(30):
        current_date = start_date + timedelta(days=i)
        simple_date = f"{current_date.month}/{current_date.day}" 
        is_weekend = current_date.weekday() >= 5
        day_type = "假日" if is_weekend else "平日"
        
        display_str = f"{simple_date} ({day_type})"
        days_info[simple_date] = {"類型": day_type, "完整": display_str}
        days_display.append(display_str)

    # ------------------------------------------
    # 頁籤 1：前台 (給醫師填寫)
    # ------------------------------------------
    with tab1:
        st.subheader("登記值班意願")
        
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
            
            # 動態計算目前每個日期有多少人選
            date_counts = {d: 0 for d in days_display}
            for doc, prefs in st.session_state.preferences.items():
                for p in prefs:
                    if p in date_counts:
                        date_counts[p] += 1
            
            preferred_days = st.multiselect(
                "請勾選您【希望值班】的日期：", 
                options=days_display,
                format_func=lambda x: f"{x} ｜ 目前有 {date_counts.get(x, 0)} 人選擇",
                default=st.session_state.preferences.get(selected_doctor, [])
            )
            
            if st.button("送出意願", key="submit_btn"):
                st.session_state.preferences[selected_doctor] = preferred_days
                st.success("意願已成功送出！")

    # ------------------------------------------
    # 頁籤 2：後台 (自動排班與下載 Excel)
    # ------------------------------------------
    with tab2:
        st.subheader("執行自動排班")
        
        if st.button("產生最終班表", type="primary"):
            schedule = {date: None for date in days_info.keys()}
            regular_second_line = [doc for doc in second_line_docs if doc not in special_second_line]
            
            # 優先級 0：填入固定值班
            for idx, row in df.iterrows():
                doc_name = row['醫師姓名']
                fixed_str = str(row.get('固定值班日期', ''))
                if fixed_str.strip():
                    dates = [d.strip() for d in fixed_str.split(',')]
                    for d in dates:
                        if d in schedule:
                            schedule[d] = doc_name
            
            # 階段一：一線醫師填空
            for d, doc in schedule.items():
                if doc is None:
                    interested_first = [doc for doc in first_line_docs if doc in st.session_state.preferences and days_info[d]["完整"] in st.session_state.preferences[doc]]
                    if interested_first: schedule[d] = random.choice(interested_first)
            
            # 階段二：特定二線優先
            for d, doc in schedule.items():
                if doc is None:
                    interested_special = [doc for doc in special_second_line if doc in st.session_state.preferences and days_info[d]["完整"] in st.session_state.preferences[doc]]
                    if interested_special: schedule[d] = random.choice(interested_special)
            
            # 階段三：一般二線填補 (盡量平均分配)
            for d, doc in schedule.items():
                if doc is None:
                    if regular_second_line: 
                        schedule[d] = random.choice(regular_second_line)
                    else: 
                        schedule[d] = "待補"

            # 整理最終表格，加上 MVPN
            final_schedule_list = []
            for d, doc in schedule.items():
                display_name = doc
                if doc in mvpn_dict and mvpn_dict[doc] != "":
                    display_name = f"{doc} ({mvpn_dict[doc]})"
                
                final_schedule_list.append({
                    "日期": d,
                    "平假日": days_info[d]["類型"],
                    "值班醫師": display_name
                })
                
            df_result = pd.DataFrame(final_schedule_list)
            st.dataframe(df_result, use_container_width=True)
            
            # 將結果轉換為真實的 Excel (.xlsx) 檔案
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