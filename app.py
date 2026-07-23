import streamlit as st
import pandas as pd
import altair as alt
import base64
import os
import json
import re
import datetime

# 🔽 Firestore用のライブラリを追加
import firebase_admin
from firebase_admin import credentials
from firebase_admin import firestore

from config import all_staff_list, CATEGORY_LIST, YEAR_LIST, MONTH_LIST, FULL_DETAIL_A, FULL_DETAIL_B, FULL_DETAIL_C, FULL_DETAIL_D

# --- Firestoreの初期化 ---
# Streamlitは画面更新のたびにコードが再実行されるため、既に接続済みの場合はスキップさせます
if not firebase_admin._apps:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred)
db = firestore.client()

# 画面をワイドに使う設定
st.set_page_config(layout="wide", page_title="MITAKE 社内管理アプリ")

# 文字サイズを小さくし、上下の余白を詰めるデザイン設定
st.markdown("""
<style>
    /* 全体の文字サイズを小さくする */
    html, body, [class*="css"]  {
        font-size: 14px !important;
    }
    /* 画面上下の余白を極力なくして1画面に収める */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔑 APIキーの設定
# ==========================================
try:
    GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
except:
    GEMINI_API_KEY = "AQ.Ab8RN6LqA9XL7FdgnE5dJjaYBxFXRb2bZMCQ1fahjBd1W-MHEQ"

try:
    import google.generativeai as genai
    from PIL import Image, ImageEnhance, ImageFilter
    if GEMINI_API_KEY != "YOUR_KEY_HERE":
        genai.configure(api_key=GEMINI_API_KEY)
except ImportError:
    st.error("⚠️ AI機能を使うためのライブラリが不足しています。")

# ==========================================
# ⚡ データの読み込み＆キャッシュ（記憶）設定
# ==========================================
@st.cache_data(ttl=600)
def load_csv_data(url):
    try:
        return pd.read_csv(url)
    except FileNotFoundError:
        return pd.DataFrame()

@st.cache_data(ttl=600)
def load_excel_data(path):
    try:
        return pd.read_excel(path, sheet_name=0, header=None)
    except Exception:
        return pd.DataFrame()

# ==========================================
# 💡 共通サイドバー（グローバルフィルター）の構築
# ==========================================
with st.sidebar:
    st.title("MITAKE 管理システム")
    st.markdown("---")
    st.markdown("### 🔍 共通フィルター")
    
    global_target_year = st.selectbox(
        "対象の期", 
        YEAR_LIST, 
        index=YEAR_LIST.index("57期") if "57期" in YEAR_LIST else 0, 
        key="global_year"
    )
    global_target_month = st.selectbox(
        "対象の月度", 
        ["通期（全月合計）"] + MONTH_LIST, 
        index=0, 
        key="global_month"
    )
    st.markdown("---")
    st.info("💡 ここで設定した「期」と「月度」が、各集計タブ（ダッシュボード・会社全体実績など）に自動で連動します。")


# 💡 同じ名前（AI読取など）の項目が複数あっても絶対に消去せず保持するロジック
def merge_details(full_template, current_data):
    if not current_data:
        return [item.copy() for item in full_template]
        
    cur_dict = {}
    for item in current_data:
        koshu = item.get("工種名", "")
        if koshu not in cur_dict:
            cur_dict[koshu] = []
        cur_dict[koshu].append(item)
        
    merged = []
    for f_item in full_template:
        koshu = f_item["工種名"]
        if koshu in cur_dict and len(cur_dict[koshu]) > 0:
            item = cur_dict[koshu].pop(0)
            if "業者管理番号" not in item: item["業者管理番号"] = ""
            if "業者・工種名" not in item: item["業者・工種名"] = item.get("業者名", "")
            if "協力業者支払" not in item: item["協力業者支払"] = 0
            if "完了金額" not in item: item["完了金額"] = 0
            if "実行予算" not in item: item["実行予算"] = 0
            merged.append(item)
        else:
            merged.append(f_item.copy())
            
    # テンプレートにない行（追加されたAI読取データなど）をすべて消さずに末尾に追加する
    for koshu, items in cur_dict.items():
        for item in items:
            if "業者管理番号" not in item: item["業者管理番号"] = ""
            if "業者・工種名" not in item: item["業者・工種名"] = item.get("業者名", "")
            if "協力業者支払" not in item: item["協力業者支払"] = 0
            if "完了金額" not in item: item["完了金額"] = 0
            if "実行予算" not in item: item["実行予算"] = 0
            merged.append(item)
            
    return merged

# ==========================================
# 💡 50行業者マスタ管理
# ==========================================
file_contractor = "data_contractor.csv"
if os.path.exists(file_contractor):
    try:
        df_contractor_master = pd.read_csv(file_contractor)
        df_contractor_master = df_contractor_master.fillna("")
        if "業者管理番号" not in df_contractor_master.columns or len(df_contractor_master) < 50:
            raise Exception("Old Format")
        df_contractor_master["区分"] = df_contractor_master["区分"].replace({"A": "内", "B": "設", "C": "電", "D": "P"})
    except:
        df_contractor_master = pd.DataFrame({"区分": [""] * 50, "業者管理番号": [f"業{i:03d}" for i in range(1, 51)], "業者・工種名": [""] * 50})
        df_contractor_master.to_csv(file_contractor, index=False)
else:
    df_contractor_master = pd.DataFrame({"区分": [""] * 50, "業者管理番号": [f"業{i:03d}" for i in range(1, 51)], "業者・工種名": [""] * 50})
    df_contractor_master.to_csv(file_contractor, index=False)

def backup_uploaded_file(file_bytes, original_filename, prefix="backup"):
    storage_dir = "storage_archive"
    os.makedirs(storage_dir, exist_ok=True)
    current_ymd = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    ext = original_filename.split('.')[-1] if '.' in original_filename else 'xlsx'
    file_path = os.path.join(storage_dir, f"{prefix}_{current_ymd}.{ext}")
    with open(file_path, "wb") as f:
        f.write(file_bytes)
    return file_path

# ==========================================
# 💡 物件データ管理 (Firestore版)
# ==========================================
def get_default_project_data():
    return {
        "client_name": "未設定", "project_number": "", "project_year": "57期",  
        "project_month": "9月", "status": "進行中",        
        "eval_memo": "", "budget_memo": "", "materials_memo": "",
        "calc_image": "",
        "受注金額": 0,
        "project_dept": "内",
        "bunrui_1": "", "bunrui_2": "", "motouke_shitauke": "",
        "detail_a": [item.copy() for item in FULL_DETAIL_A], 
        "detail_b": [item.copy() for item in FULL_DETAIL_B],
        "detail_c": [item.copy() for item in FULL_DETAIL_C], 
        "detail_d": [item.copy() for item in FULL_DETAIL_D],
        "detail_request": [],
        "staff_setsubi": [],
        "staff_naisou": [],
        "staff_denki": [],
        "staff_pm": []
    }

def load_projects():
    year_to_term = {"2022年": "53期", "2023年": "54期", "2024年": "55期", "2025年": "56期", "2026年": "57期", "2027年": "58期", "2028年": "59期", "2029年": "60期", "2030年": "61期"}
    valid_data = {}
    
    try:
        # 🔽 Firestoreから全データを取得
        docs = db.collection("projects").stream()
        for doc in docs:
            p_name = doc.id
            p_data = doc.to_dict()
            
            if not isinstance(p_data, dict): continue
            if "project_number" not in p_data: p_data["project_number"] = ""
            if "client_name" not in p_data: p_data["client_name"] = "未設定"
            if "project_month" not in p_data: p_data["project_month"] = "9月"
            if "status" not in p_data: p_data["status"] = "進行中"
            if "materials_memo" not in p_data: p_data["materials_memo"] = ""
            if "calc_image" not in p_data: p_data["calc_image"] = ""
            if "受注金額" not in p_data: p_data["受注金額"] = 0
            if "project_dept" not in p_data: p_data["project_dept"] = "内"
            if "bunrui_1" not in p_data: p_data["bunrui_1"] = p_data.get("category", "")
            if "bunrui_2" not in p_data: p_data["bunrui_2"] = ""
            if "motouke_shitauke" not in p_data: p_data["motouke_shitauke"] = ""
            if "project_year" not in p_data: p_data["project_year"] = "57期"
            elif p_data["project_year"] in year_to_term: p_data["project_year"] = year_to_term[p_data["project_year"]]
            
            valid_data[p_name] = p_data
    except Exception as e:
        print(f"Firestore読み込みエラー: {e}")

    # 初期データがない場合のダミー生成
    if not valid_data:
        valid_data = {"渋谷カフェA": get_default_project_data()}
        valid_data["渋谷カフェA"]["client_name"] = "株式会社OMC"
        valid_data["渋谷カフェA"]["project_number"] = "物001"
        valid_data["渋谷カフェA"]["project_year"] = "57期"
        valid_data["渋谷カフェA"]["project_month"] = "9月"
        
    return valid_data

def save_projects(data):
    # 🔽 辞書データをFirestoreに上書き保存
    for p_name, p_data in data.items():
        db.collection("projects").document(p_name).set(p_data)

projects_db = load_projects()

keys_to_delete = []
for k, v in projects_db.items():
    if not isinstance(v, dict): continue
    c_name = str(v.get("client_name", ""))
    p_month = str(v.get("project_month", ""))
    if "合計" in c_name or p_month == "通期合計" or c_name == "未設定 (自動追加)":
        keys_to_delete.append(k)

if keys_to_delete:
    for k in keys_to_delete: 
        del projects_db[k]
        # 🔽 Firestore上の無効データも直接削除
        db.collection("projects").document(k).delete() 

# 💡 【数値読取の鉄壁フィルター】
def safe_num(v):
    try:
        if v is not None and str(v).strip() != "":
            clean_v = str(v).translate(str.maketrans('０１２３４５６７８９', '0123456789'))
            clean_v = re.sub(r'[^\d.-]', '', clean_v)
            if not clean_v: return 0.0
            return float(clean_v)
        else: return 0.0
    except: return 0.0

def get_hours_for_staff(df, staff_name, current_project):
    if df is None or df.empty: return "0.0"
    df.columns = [c.strip() for c in df.columns]
    target_col = "質問1：担当者名"
    if target_col not in df.columns:
        for col in df.columns:
            if "担当者" in col:
                target_col = col; break
    if target_col not in df.columns or "打刻区分" not in df.columns or "物件名" not in df.columns: return "0.0"
    df_filtered = df[(df[target_col] == staff_name) & (df["物件名"] == current_project)]
    df_filtered = df_filtered.sort_values(by="Timestamp")
    total = 0.0
    st_time = None
    for _, row in df_filtered.iterrows():
        try:
            c_time = pd.to_datetime(row["Timestamp"])
            status = str(row["打刻区分"]).strip()
            if status == "出勤": st_time = c_time
            elif status == "退勤" and st_time is not None:
                total += (c_time - st_time).total_seconds() / 3600.0
                st_time = None
        except: pass
    return f"{total:.1f}"

def calc_hours_from_time_str(time_str):
    try:
        time_str = str(time_str).strip()
        if not time_str or time_str.lower() in ["nan", "nat", "none", ""]: return 0.0
        try: return float(time_str)
        except ValueError: pass
        time_str = time_str.translate(str.maketrans('０１２３４５６７８９：〜～ー−', '0123456789:----'))
        time_str = time_str.replace("~", "-")
        time_str = re.sub(r'[^0-9:\-]', '', time_str)
        if not time_str: return 0.0
        parts = time_str.split("-")
        if len(parts) == 2:
            start_t = datetime.datetime.strptime(parts[0].strip(), "%H:%M")
            end_t = datetime.datetime.strptime(parts[1].strip(), "%H:%M")
            diff = (end_t - start_t).total_seconds() / 3600.0
            break_time = 1.0 if diff >= 6.0 else 0.0
            return max(0.0, diff - break_time)
        return 0.0 
    except: return 0.0 

# ==========================================
# 🚀 複数ANDPADエクセルファイルの「累積・合算」解析エンジン
# ==========================================
ANDPAD_ACC_DIR = "uploaded_images/andpad_accumulated"
os.makedirs(ANDPAD_ACC_DIR, exist_ok=True)

def parse_single_andpad_excel(path):
    try:
        df = pd.read_excel(path, sheet_name='予定表', header=None)
        parsed_data = {}
        data_values = df.values 
        rows_count = len(data_values)
        current_staff = None
        for i in range(4, rows_count):
            row = data_values[i]
            if pd.notna(row[2]) and str(row[2]) != '氏名':
                raw_staff_name = str(row[2])
                current_staff = re.sub(r'\s+', '', raw_staff_name) 
                if current_staff not in parsed_data: parsed_data[current_staff] = []
            if len(row) > 4 and str(row[4]) == '案件名' and current_staff:
                projects = row[5:]
                times = data_values[i+1][5:] if (i+1) < rows_count else []
                for col_idx, project in enumerate(projects):
                    if pd.notna(project):
                        p_str = re.sub(r'\s+', '', str(project)) 
                        raw_name = str(project).strip()
                        t_str = str(times[col_idx]) if col_idx < len(times) else ""
                        hours = calc_hours_from_time_str(t_str)
                        if hours > 0:
                            parsed_data[current_staff].append({"proj": p_str, "raw_name": raw_name, "hours": hours})
        return parsed_data
    except: return {}

def get_accumulated_files_tuple():
    if not os.path.exists(ANDPAD_ACC_DIR): return ()
    files = sorted([os.path.join(ANDPAD_ACC_DIR, f) for f in os.listdir(ANDPAD_ACC_DIR) if f.endswith(('.xlsx', '.xls'))])
    return tuple((f, os.path.getmtime(f)) for f in files)

@st.cache_data(show_spinner=False)
def get_accumulated_andpad_dict(file_mtime_tuple):
    master_parsed = {}
    for path, _ in file_mtime_tuple:
        single_data = parse_single_andpad_excel(path)
        for staff, items in single_data.items():
            if staff not in master_parsed: master_parsed[staff] = []
            master_parsed[staff].extend(items)
    return master_parsed

accumulated_files_tuple = get_accumulated_files_tuple()
parsed_andpad_data = get_accumulated_andpad_dict(accumulated_files_tuple)

def get_hours_from_andpad(staff_name, current_project, project_number=""):
    if not parsed_andpad_data: return 0.0
    safe_staff = re.sub(r'\s+', '', str(staff_name))
    cp = re.sub(r'\s+', '', str(current_project))
    pn = re.sub(r'\s+', '', str(project_number))
    total = 0.0
    for excel_staff, items in parsed_andpad_data.items():
        is_staff_match = False
        if safe_staff in excel_staff or excel_staff in safe_staff: is_staff_match = True
        elif len(safe_staff) >= 2 and safe_staff[:2] in excel_staff: is_staff_match = True
        if is_staff_match:
            for item in items:
                c_proj = item["proj"]
                if not c_proj: continue
                is_match = False
                if pn and (pn in c_proj): is_match = True
                elif cp and (cp in c_proj): is_match = True
                elif len(c_proj) >= 2 and (c_proj in cp): is_match = True
                if is_match: total += item["hours"]
    return total

def extract_projects_and_hours_from_dict():
    if not parsed_andpad_data: return []
    proj_hours = {}
    for staff, items in parsed_andpad_data.items():
        for item in items:
            raw_name = item["raw_name"]
            if raw_name: proj_hours[raw_name] = proj_hours.get(raw_name, 0.0) + item["hours"]
    display_list = []
    for proj, hrs in sorted(proj_hours.items()):
        display_list.append(f"{proj}  👉  {hrs:.1f} h")
    return display_list

# ==========================================
# 💡 タブ作成
# ==========================================
tab_dash, tab1, tab2, tab3, tab4 = st.tabs(["🏠 ダッシュボード", "📊 会社全体・部署別", "🏗️ 物件別予算管理", "📸 AI書類自動処理", "📈 顧客別分析"])

# --- ダッシュボードタブの中身 ---
with tab_dash:
    st.markdown("## 🏢 会社全体 プロジェクト・ダッシュボード")
    st.info(f"💡 現在、左のサイドバーで **【{global_target_year}】** が選択されています。")
    st.divider()

    dash_sales_a = dash_sales_b = dash_sales_c = dash_sales_d = 0
    dash_cost_a = dash_cost_b = dash_cost_c = dash_cost_d = dash_cost_req = 0
    total_sales = 0
    total_cost = 0
    
    for p_name, p_data in projects_db.items():
        if not isinstance(p_data, dict): continue
        if p_data.get("project_year", "57期") != global_target_year: continue
        
        val_a = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_a", []))
        val_b = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_b", []))
        val_c = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_c", []))
        val_d = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_d", []))
        
        order_amt = safe_num(p_data.get("受注金額", 0))
        p_dept = p_data.get("project_dept", "内")
        
        if order_amt > 0:
            if p_dept == "内":   cur_sa = order_amt; cur_sb = 0; cur_sc = 0; cur_sd = 0
            elif p_dept == "設": cur_sa = 0; cur_sb = order_amt; cur_sc = 0; cur_sd = 0
            elif p_dept == "電": cur_sa = 0; cur_sb = 0; cur_sc = order_amt; cur_sd = 0
            elif p_dept == "P":  cur_sa = 0; cur_sb = 0; cur_sc = 0; cur_sd = order_amt
            else:                cur_sa = val_a; cur_sb = val_b; cur_sc = val_c; cur_sd = val_d
            total_sales += order_amt
        else:
            cur_sa = val_a; cur_sb = val_b; cur_sc = val_c; cur_sd = val_d
            total_sales += (val_a + val_b + val_c + val_d)
            
        dash_sales_a += cur_sa; dash_sales_b += cur_sb; dash_sales_c += cur_sc; dash_sales_d += cur_sd
        
        cost_a = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_a", []))
        cost_b = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_b", []))
        cost_c = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_c", []))
        cost_d = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_d", []))
        cost_req = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_request", []))
        
        dash_cost_a += cost_a; dash_cost_b += cost_b; dash_cost_c += cost_c; dash_cost_d += cost_d
        dash_cost_req += cost_req
        total_cost += (cost_a + cost_b + cost_c + cost_d + cost_req)

    gross_profit = total_sales - total_cost
    profit_margin = (gross_profit / total_sales) * 100 if total_sales > 0 else 0

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1: st.metric(label="💰 全体受注金額", value=f"¥{total_sales:,.0f}")
    with col2: st.metric(label="💸 全体発注金額", value=f"¥{total_cost:,.0f}")
    with col3: st.metric(label="💵 会社全体 純利益", value=f"¥{gross_profit:,.0f}", delta="粗利額", delta_color="normal")
    with col4: st.metric(label="📊 会社全体 粗利率", value=f"{profit_margin:.1f}%")
    with col5:
        if total_sales > 0 and profit_margin < 10: st.metric(label="⚠️ アラート", value="警告", delta="全体粗利率が10%未満です", delta_color="inverse")
        else: st.metric(label="✅ ステータス", value="正常", delta="利益率クリア", delta_color="normal")

    st.markdown("<br>", unsafe_allow_html=True)
    col_graph1, col_graph2 = st.columns([1, 1])
    
    with col_graph1:
        st.markdown("#### 📈 全期ごとの売上・利益推移")
        term_summary = {year: {"売上": 0.0, "利益": 0.0} for year in YEAR_LIST}
        for p_name, p_data in projects_db.items():
            if not isinstance(p_data, dict): continue
            p_year = p_data.get("project_year", "57期")
            if p_year in term_summary:
                val_a = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_a", []))
                val_b = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_b", []))
                val_c = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_c", []))
                val_d = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_d", []))
                order_amt = safe_num(p_data.get("受注金額", 0))
                p_sales = order_amt if order_amt > 0 else (val_a + val_b + val_c + val_d)
                p_cost = sum(safe_num(x.get("協力業者支払")) for c in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"] for x in p_data.get(c, []))
                term_summary[p_year]["売上"] += p_sales
                term_summary[p_year]["利益"] += (p_sales - p_cost)
       
        df_term_graph = pd.DataFrame([{"期": year, "売上": term_summary[year]["売上"], "利益": term_summary[year]["利益"]} for year in YEAR_LIST]).set_index("期")
        st.line_chart(df_term_graph)

    with col_graph2:
        st.markdown("#### 🍩 カテゴリ別 売上構成比")
        df_pie = pd.DataFrame({"カテゴリ": ["内装", "設備", "電気", "厨房"], "売上": [dash_sales_a, dash_sales_b, dash_sales_c, dash_sales_d]})
        if df_pie["売上"].sum() == 0: st.info("まだ今期の売上データがありません")
        else:
            pie_chart = alt.Chart(df_pie).mark_arc(innerRadius=50).encode(theta=alt.Theta(field="売上", type="quantitative"), color=alt.Color(field="カテゴリ", type="nominal"), tooltip=["カテゴリ", "売上"]).properties(height=300)
            st.altair_chart(pie_chart, use_container_width=True)

    st.markdown("---")
    st.markdown("### 🔍 個別顧客・商社協力業者のクイック抽出コーナー")
    dash_sel_c1, dash_sel_c2 = st.columns(2)
    
    with dash_sel_c1:
        st.markdown("#### 🏢 顧客を選んで今期実績を表示")
        dash_all_clients = sorted(list(set([v.get("client_name", "未設定") for v in projects_db.values() if isinstance(v, dict)])))
        dash_all_clients = [c for c in dash_all_clients if c != "未設定" and "合計" not in c and "修理" not in c and "スキップ" not in c]
        selected_dash_client = st.selectbox("確認したい顧客（会社名）を選択してください", ["(未選択)"] + dash_all_clients, key="dash_client_sb_v24")
        
        if selected_dash_client != "(未選択)":
            c_sales = c_cost = 0
            for p_name, p_data in projects_db.items():
                if p_data.get("project_year") == global_target_year and p_data.get("client_name") == selected_dash_client:
                    c_sales += safe_num(p_data.get("受注金額", 0))
                    c_cost += sum(safe_num(x.get("協力業者支払")) for c in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"] for x in p_data.get(c, []))
            c_profit = c_sales - c_cost
            c_rate = (c_profit / c_sales) * 100 if c_sales > 0 else 0.0
            met_c1, met_c2, met_c3 = st.columns(3)
            met_c1.metric("受注総額", f"¥{c_sales:,.0f}")
            met_c2.metric("純利益額", f"¥{c_profit:,.0f}")
            met_c3.metric("顧客別粗利率", f"{c_rate:.1f}%")

    with dash_sel_c2:
        st.markdown("#### 🤝 商社・協力業者を選んで支払内訳を表示")
        all_contractor_names = sorted(list(set([str(x.get("業者・工種名", "")).strip() for v in projects_db.values() if isinstance(v, dict) for c in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"] for x in v.get(c, []) if str(x.get("業者・工種名", "")).strip() != ""])))
        selected_dash_contractor = st.selectbox("確認したい業者・工種名を選択してください", ["(未選択)"] + all_contractor_names, key="dash_contractor_sb_v24")
        
        if selected_dash_contractor != "(未選択)":
            con_rows = []
            for p_name, p_data in projects_db.items():
                if p_data.get("project_year") == global_target_year:
                    for cat in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"]:
                        for item in p_data.get(cat, []):
                            if str(item.get("業者・工種名", "")).strip() == selected_dash_contractor:
                                con_rows.append({"物件名": p_name, "月度": p_data.get("project_month", "9月"), "実行予算(円)": safe_num(item.get("実行予算", 0)), "支払金額(円)": safe_num(item.get("協力業者支払", 0))})
            if con_rows:
                df_con_res = pd.DataFrame(con_rows)
                st.metric("今期の総支払額", f"¥{df_con_res['支払金額(円)'].sum():,.0f}")
                st.dataframe(df_con_res, use_container_width=True, hide_index=True)
            else: st.info("今期はこの業者さんへの支払データが登録されていません。")

# --- タブ1：会社全体・部署別 ---
with tab1:
    summary_placeholder = st.container()
    st.markdown("---")
    
    sales_a = sales_b = sales_c = sales_d = 0
    dept_cost_a = dept_cost_b = dept_cost_c = dept_cost_d = dept_cost_req = 0
    contractor_costs_by_name = {}
    project_sales_rows = []
    monthly_data_dict = {m: {"売上": 0.0, "支払": 0.0} for m in MONTH_LIST}

    for p_name, p_data in projects_db.items():
        if not isinstance(p_data, dict): continue
        if p_data.get("project_year", "57期") != global_target_year: continue
            
        c_name = p_data.get("client_name", "未設定")
        p_num = p_data.get("project_number", "")
        p_month = p_data.get("project_month", "9月")
        order_amount = safe_num(p_data.get("受注金額", 0))
        p_dept = p_data.get("project_dept", "内")
        
        val_a = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_a", []))
        val_b = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_b", []))
        val_c = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_c", []))
        val_d = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_d", []))
        
        cost_a = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_a", []))
        cost_b = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_b", []))
        cost_c = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_c", []))
        cost_d = sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_d", []))
        
        if order_amount > 0: p_sales_total = order_amount
        else: p_sales_total = val_a + val_b + val_c + val_d
        
        val_sum = val_a + val_b + val_c + val_d
        cost_sum = cost_a + cost_b + cost_c + cost_d
        
        if val_sum > 0:
            cur_sa = p_sales_total * (val_a / val_sum)
            cur_sb = p_sales_total * (val_b / val_sum)
            cur_sc = p_sales_total * (val_c / val_sum)
            cur_sd = p_sales_total * (val_d / val_sum)
        elif cost_sum > 0:
            cur_sa = p_sales_total * (cost_a / cost_sum)
            cur_sb = p_sales_total * (cost_b / cost_sum)
            cur_sc = p_sales_total * (cost_c / cost_sum)
            cur_sd = p_sales_total * (cost_d / cost_sum)
        else:
            if order_amount > 0:
                if p_dept == "内":   cur_sa = order_amount; cur_sb = 0; cur_sc = 0; cur_sd = 0
                elif p_dept == "設": cur_sa = 0; cur_sb = order_amount; cur_sc = 0; cur_sd = 0
                elif p_dept == "電": cur_sa = 0; cur_sb = 0; cur_sc = order_amount; cur_sd = 0
                elif p_dept == "P":  cur_sa = 0; cur_sb = 0; cur_sc = 0; cur_sd = order_amount
                else:                cur_sa = val_a; cur_sb = val_b; cur_sc = val_c; cur_sd = val_d
            else:
                cur_sa = val_a; cur_sb = val_b; cur_sc = val_c; cur_sd = val_d
        
        if p_month in monthly_data_dict:
            monthly_data_dict[p_month]["売上"] += p_sales_total
            for cat in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"]:
                for item in p_data.get(cat, []):
                    monthly_data_dict[p_month]["支払"] += safe_num(item.get("協力業者支払"))
                    
        if global_target_month != "通期（全月合計）" and p_month != global_target_month: continue

        p_total_hours = 0.0
        for dept in ["staff_setsubi", "staff_naisou", "staff_denki", "staff_pm"]:
            for staff in p_data.get(dept, []):
                p_total_hours += get_hours_from_andpad(staff, p_name, p_num)
            
        sales_a += cur_sa; sales_b += cur_sb; sales_c += cur_sc; sales_d += cur_sd
        
        dept_cost_a += cost_a
        dept_cost_b += cost_b
        dept_cost_c += cost_c
        dept_cost_d += cost_d
        dept_cost_req += sum(safe_num(x.get("協力業者支払")) for x in p_data.get("detail_request", []))
        
        p_cost_total = 0.0
        for cat in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"]:
            for item in p_data.get(cat, []):
                g_name = str(item.get("業者・工種名", "")).strip()
                c_val = safe_num(item.get("協力業者支払"))
                p_cost_total += c_val
                if g_name and g_name.lower() != "none": contractor_costs_by_name[g_name] = contractor_costs_by_name.get(g_name, 0.0) + c_val

        p_profit = p_sales_total - p_cost_total
        
        project_sales_rows.append({
            "物件番号": p_num, "顧客名": c_name, "物件名": p_name, "月度": p_month,
            "売上高 合計 (円)": p_sales_total, "協力業者支払 合計 (円)": p_cost_total, "売上純利益 (円)": p_profit,
            "総投入工数 (h)": p_total_hours
        })

    monthly_rows = []
    for m_str in MONTH_LIST:
        m_sales = monthly_data_dict[m_str]["売上"]
        m_cost = monthly_data_dict[m_str]["支払"]
        m_profit = m_sales - m_cost
        monthly_rows.append({"月度": m_str, "顧客売上高 (円)": m_sales, "商社・協力業者支払 (円)": m_cost, "売上純利益 (円)": m_profit})
    
    df_monthly_calc = pd.DataFrame(monthly_rows)

    sales_grid_rows = []
    for i in range(300):
        if i < len(project_sales_rows): sales_grid_rows.append(project_sales_rows[i])
        else: sales_grid_rows.append({"物件番号": "", "顧客名": "", "物件名": "", "月度": "", "売上高 合計 (円)": 0.0, "協力業者支払 合計 (円)": 0.0, "売上純利益 (円)": 0.0, "総投入工数 (h)": 0.0})
    df_project_sales = pd.DataFrame(sales_grid_rows)

    df_contractor_display = df_contractor_master.copy()
    df_contractor_display["支払金額 (円)"] = df_contractor_display["業者・工種名"].apply(lambda n: contractor_costs_by_name.get(str(n).strip(), 0.0))
    
    master_names = set(df_contractor_display["業者・工種名"].astype(str).str.strip().tolist())
    unmastered_rows = []
    for name, cost in contractor_costs_by_name.items():
        if name and name not in master_names and cost > 0:
            unmastered_rows.append({"区分": "AI/新規", "業者・工種名": name, "支払金額 (円)": cost})
            
    if unmastered_rows:
        df_unmastered = pd.DataFrame(unmastered_rows)
        df_contractor_display = pd.concat([df_unmastered, df_contractor_display], ignore_index=True)

    if global_target_month == "通期（全月合計）":
        calc_sales = df_monthly_calc["顧客売上高 (円)"].sum()
        calc_cost = df_monthly_calc["商社・協力業者支払 (円)"].sum()
        calc_profit = df_monthly_calc["売上純利益 (円)"].sum()
    else:
        target_row = df_monthly_calc[df_monthly_calc["月度"] == global_target_month]
        calc_sales = target_row["顧客売上高 (円)"].sum() if not target_row.empty else 0
        calc_cost = target_row["商社・協力業者支払 (円)"].sum() if not target_row.empty else 0
        calc_profit = target_row["売上純利益 (円)"].sum() if not target_row.empty else 0
        
    calc_rate = (calc_profit / calc_sales) * 100 if calc_sales > 0 else 0.0

    with summary_placeholder:
        st.header(f"【{global_target_year} {global_target_month}】 会社全体 売上サマリー")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("売上金額", f"{calc_sales:,.0f} 円")
        col2.metric("支払金額", f"{calc_cost:,.0f} 円")
        col3.metric("売上純利益", f"{calc_profit:,.0f} 円")
        col4.metric("粗利率", f"{calc_rate:.1f} %")

    t1_col1, t1_col2 = st.columns(2)
    
    with t1_col1:
        st.subheader(f"📅 【{global_target_year}】 月次集計表（※全期固定表示）")
        st.dataframe(df_monthly_calc, use_container_width=True, hide_index=True, height=220)
        
        df_monthly_melted = df_monthly_calc.melt("月度", value_vars=["顧客売上高 (円)", "商社・協力業者支払 (円)", "売上純利益 (円)"], var_name="項目", value_name="金額")
        monthly_chart = alt.Chart(df_monthly_melted).mark_bar().encode(x=alt.X('月度:N', sort=MONTH_LIST, title=None), y=alt.Y('金額:Q', title=None), color='項目:N', xOffset='項目:N').properties(height=200)
        st.altair_chart(monthly_chart, use_container_width=True)
        
        st.markdown("---")
        st.subheader(f"📊 物件（顧客）別 売上・利益集計表 ({global_target_month}連動)")
        
        t1_client_list = sorted(list(set([str(x) for x in df_project_sales["顧客名"].tolist() if str(x) != ""])))
        filter_client = st.selectbox("🏢 顧客名で絞り込む:", ["(すべての顧客)"] + t1_client_list, key="filter_client_sb_v24")
        df_project_sales_filtered = df_project_sales.copy()
        if filter_client != "(すべての顧客)": df_project_sales_filtered = df_project_sales[df_project_sales["顧客名"] == filter_client]
        st.dataframe(df_project_sales_filtered, use_container_width=True, hide_index=True, height=380)
        
        st.markdown("---")
        st.subheader(f"👤 担当者別 携わっている物件・就業時間一覧 ({global_target_month}連動)")
        st.info("※担当表（Excel）のアップロード機能は、「📸 AI 書類自動処理」タブに統合されました。")
        
        selected_staff = st.selectbox("👨‍💼 情報を確認したい担当者を選択してください", all_staff_list, key="staff_chk_sel_v24")
        
        staff_projects_rows = []
        for p_name, p_data in projects_db.items():
            if not isinstance(p_data, dict): continue
            if p_data.get("project_year", "57期") != global_target_year: continue
            if global_target_month != "通期（全月合計）" and p_data.get("project_month", "9月") != global_target_month: continue
            
            is_member = False
            for dept in ["staff_setsubi", "staff_naisou", "staff_denki", "staff_pm"]:
                if selected_staff in p_data.get(dept, []):
                    is_member = True
                    break
                    
            p_num = p_data.get("project_number", "")
            staff_hours = get_hours_from_andpad(selected_staff, p_name, p_num)
                    
            if is_member or staff_hours > 0:
                staff_projects_rows.append({
                    "物件番号": p_num, "物件名": p_name, "顧客名": p_data.get("client_name", "未設定"),
                    "所属部署での登録枠": "✅ メンバー登録あり" if is_member else "❌ 登録なし（打刻実績のみ）", "その物件での累計就業時間": f"{staff_hours:.1f} h"
                })
                
        staff_display_rows = []
        for i in range(300):
            if i < len(staff_projects_rows): staff_display_rows.append(staff_projects_rows[i])
            else: staff_display_rows.append({"物件番号": "", "物件名": "", "顧客名": "", "所属部署での登録枠": "", "その物件での累計就業時間": ""})
        st.dataframe(pd.DataFrame(staff_display_rows), use_container_width=True, hide_index=True, height=380)

        # --- 物件別の総投入時間グラフ ---
        st.markdown("---")
        st.subheader(f"⏱️ 物件別の総投入工数（就業時間）可視化 ({global_target_month}連動)")
        
        df_proj_hours = pd.DataFrame(project_sales_rows)
        if not df_proj_hours.empty and "総投入工数 (h)" in df_proj_hours.columns:
            df_hours_filtered = df_proj_hours[df_proj_hours["総投入工数 (h)"] > 0]
            if not df_hours_filtered.empty:
                hour_chart = alt.Chart(df_hours_filtered).mark_bar(color='#4A90E2').encode(
                    x=alt.X('総投入工数 (h):Q', title='合計時間 (h)'),
                    y=alt.Y('物件名:N', sort='-x', title='物件名'),
                    tooltip=['顧客名', '物件名', '総投入工数 (h)']
                ).properties(height=max(200, len(df_hours_filtered) * 40))
                st.altair_chart(hour_chart, use_container_width=True)
            else:
                st.info("対象の月度において、計上されている就業時間データはありません。")

    with t1_col2:
        st.subheader(f"👥 部署別実績 ({global_target_month}連動)")
        sort_option = st.radio("表示順の変更:", ["デフォルト (固定)", "売上高が高い順", "利益が高い順"], horizontal=True, key="sort_opt_v24")
        
        df_dept_final = pd.DataFrame({
            "部署": ["第2管理部（内装）―内", "第1管理部（設備）―設", "第3管理部（電気）―電", "PM室（厨房等）―P", "🆕 未振分（AI読取）"],
            "売上高（円）": [sales_a, sales_b, sales_c, sales_d, 0],
            "商社協力業社支払（円）": [dept_cost_a, dept_cost_b, dept_cost_c, dept_cost_d, dept_cost_req],
        })
        df_dept_final["売上純利益（円）"] = df_dept_final["売上高（円）"] - df_dept_final["商社協力業社支払（円）"]
        
        if sort_option == "売上高が高い順": df_dept_final = df_dept_final.sort_values(by="売上高（円）", ascending=False)
        elif sort_option == "利益が高い順": df_dept_final = df_dept_final.sort_values(by="売上純利益（円）", ascending=False)
        st.dataframe(df_dept_final, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.subheader(f"📅 商社、協力業社支払集計表 ({global_target_month}連動)")
        
        t1_contractor_list = sorted(list(set([str(x) for x in df_contractor_display["業者・工種名"].tolist() if str(x) != ""])))
        filter_contractor = st.selectbox("🤝 業者名で絞り込む:", ["(すべての商社・協力業者)"] + t1_contractor_list, key="filter_contractor_sb_v24")
        df_contractor_display_filtered = df_contractor_display.copy()
        if filter_contractor != "(すべての商社・協力業者)": df_contractor_display_filtered = df_contractor_display[df_contractor_display["業者・工種名"] == filter_contractor]
            
        st.dataframe(
            df_contractor_display_filtered.drop(columns=["業者管理番号"], errors='ignore'), use_container_width=True, hide_index=True, height=920,
            column_config={"区分": st.column_config.TextColumn("区分"), "業者・工種名": st.column_config.TextColumn("業者・工種名"), "支払金額 (円)": st.column_config.NumberColumn("支払金額 (円)")}
        )

# --- タブ2：物件別予算管理 ---
with tab2:
    st.header("物件別予算管理")
    
    active_projects = [k for k, v in projects_db.items() if isinstance(v, dict) and v.get("status") != "アーカイブ（完了）"]
    archived_projects = [k for k, v in projects_db.items() if isinstance(v, dict) and v.get("status") == "アーカイブ（完了）"]
    
    # 💡 【バグ修正】タブによるセレクトボックスの干渉を防ぐため、ラジオボタンでモードを切り替える設計に変更しました
    view_mode = st.radio("📂 表示する物件のステータスを選択", ["🏗️ 進行中の物件", "📦 完了済み物件（アーカイブ）"], horizontal=True)
    
    selected_project = None
    if view_mode == "🏗️ 進行中の物件":
        if not active_projects:
            st.info("現在進行中の物件はありません。")
        else:
            selected_project = st.selectbox("📂 編集する進行中物件を選択", active_projects, key="select_proj_active_sub")
    else:
        if not archived_projects:
            st.info("現在完了済みの物件（アーカイブ）はありません。")
        else:
            selected_project = st.selectbox("📂 確認・編集する完了済み物件を選択", archived_projects, key="select_proj_archived_sub")

    if selected_project:
        st.markdown("---")
        cur_data = projects_db[selected_project]
        cur_data["detail_a"] = merge_details(FULL_DETAIL_A, cur_data.get("detail_a", []))
        cur_data["detail_b"] = merge_details(FULL_DETAIL_B, cur_data.get("detail_b", []))
        cur_data["detail_c"] = merge_details(FULL_DETAIL_C, cur_data.get("detail_c", []))
        cur_data["detail_d"] = merge_details(FULL_DETAIL_D, cur_data.get("detail_d", []))
        
        # 💡改善ポイント: 保存ボタンを画面上部に配置してアクセスしやすくする
        action_col1, action_col2 = st.columns([3, 1])
        with action_col1:
            st.subheader(f"選択中: {selected_project}")
        with action_col2:
            st.button("💾 変更を保存する", type="primary", use_container_width=True, key=f"save_top_{selected_project}")

        with st.expander("🗑️ 物件のステータス変更・削除"):
            current_status = projects_db[selected_project].get("status", "進行中")
            if current_status != "アーカイブ（完了）":
                if st.button("📦 完了済みとして保管する", key=f"btn_archive_{selected_project}"):
                    projects_db[selected_project]["status"] = "アーカイブ（完了）"
                    save_projects(projects_db)
                    st.success("物件を完了済み（アーカイブ）へ移動しました！")
                    st.rerun()
            else:
                if st.button("🔙 進行中に戻す", key=f"btn_unarchive_{selected_project}"):
                    projects_db[selected_project]["status"] = "進行中"
                    save_projects(projects_db)
                    st.success("進行中に戻しました！")
                    st.rerun()
            st.markdown("---")
            if st.button("🚨 完全に削除する", type="primary", key=f"btn_delete_fully_{selected_project}"):
                if selected_project in projects_db:
                    del projects_db[selected_project]
                    # 🔽 Firestoreからも直接削除
                    db.collection("projects").document(selected_project).delete()
                st.success("物件を完全に削除しました！")
                st.rerun()

        # 💡改善ポイント: 基本情報をコンテナで囲む
        with st.container(border=True):
            st.markdown("#### 基本情報")
            m_col1, m_col2, m_col3, m_col3_dept, m_col4, m_col5 = st.columns([1.2, 1, 1.2, 0.6, 0.8, 1])
            with m_col1: in_client = st.text_input("顧客名", value=cur_data.get("client_name", "未設定"), key=f"client_{selected_project}_v24")
            with m_col2: in_proj_num = st.text_input("物件番号", value=cur_data.get("project_number", ""), key=f"num_{selected_project}_v24")
            with m_col3: in_order_amount = st.number_input("受注金額", value=int(cur_data.get("受注金額", 0)), key=f"order_{selected_project}_v24")
            
            saved_dept = cur_data.get("project_dept", "内")
            if saved_dept not in ["内", "設", "電", "P"]: saved_dept = "内"
            with m_col3_dept: in_project_dept = st.selectbox("担当系統", ["内", "設", "電", "P"], index=["内", "設", "電", "P"].index(saved_dept), key=f"dept_{selected_project}_v24")
            
            safe_month = cur_data.get("project_month", "9月")
            if safe_month not in MONTH_LIST: safe_month = "9月"
            
            with m_col4:
                in_year = st.selectbox("期", YEAR_LIST, index=YEAR_LIST.index(cur_data.get("project_year", "57期")) if cur_data.get("project_year", "57期") in YEAR_LIST else 4, key=f"yr_{selected_project}_v24")
                in_month = st.selectbox("月度", MONTH_LIST, index=MONTH_LIST.index(safe_month), key=f"mnth_{selected_project}_v24")
            
            with m_col5:
                in_bunrui1 = st.text_input("分類1", value=cur_data.get("bunrui_1", ""), key=f"b1_{selected_project}_v24")
                in_bunrui2 = st.text_input("分類2", value=cur_data.get("bunrui_2", ""), key=f"b2_{selected_project}_v24")
                in_motouke = st.text_input("元請/下請", value=cur_data.get("motouke_shitauke", ""), key=f"moto_{selected_project}_v24")

        # 💡改善ポイント: 予実サマリーをコンテナで囲む
        with st.container(border=True):
            st.markdown("#### プロジェクト収支 ＆ 予実サマリー")
            
            def calc_category_totals(detail_list):
                if not detail_list: return 0, 0
                b_total = sum(safe_num(x.get("実行予算", 0)) for x in detail_list)
                p_total = sum(safe_num(x.get("協力業者支払", 0)) for x in detail_list)
                return b_total, p_total

            b_req, p_req = calc_category_totals(cur_data.get("detail_request", []))
            b_a, p_a = calc_category_totals(cur_data.get("detail_a", []))
            b_b, p_b = calc_category_totals(cur_data.get("detail_b", []))
            b_c, p_c = calc_category_totals(cur_data.get("detail_c", []))
            b_d, p_d = calc_category_totals(cur_data.get("detail_d", []))

            total_budget = b_req + b_a + b_b + b_c + b_d
            total_payment = p_req + p_a + p_b + p_c + p_d
            diff_budget = total_budget - total_payment
            consumption_rate = (total_payment / total_budget * 100) if total_budget > 0 else 0.0

            current_profit = in_order_amount - total_payment
            current_profit_margin = (current_profit / in_order_amount * 100) if in_order_amount > 0 else 0.0

            st.markdown("##### 収支状況 (受注金額ベース)")
            p_col1, p_col2, p_col3, p_col4 = st.columns(4)
            p_col1.metric("受注金額", f"¥{in_order_amount:,.0f}")
            p_col2.metric("総支払実績", f"¥{total_payment:,.0f}")
            p_col3.metric("粗利額", f"¥{current_profit:,.0f}", delta="利益", delta_color="normal")
            
            if in_order_amount > 0 and current_profit_margin < 10:
                p_col4.metric("粗利率", f"{current_profit_margin:.1f} %", delta="⚠️ 利益率低下", delta_color="inverse")
            else:
                p_col4.metric("粗利率", f"{current_profit_margin:.1f} %", delta="正常", delta_color="normal")

            st.markdown("##### 実行予算管理")
            b_col1, b_col2, b_col3, b_col4 = st.columns(4)
            b_col1.metric("総実行予算", f"¥{total_budget:,.0f}")
            b_col2.metric("総支払実績", f"¥{total_payment:,.0f}")
            b_col3.metric("予算残額", f"¥{diff_budget:,.0f}", delta=f"{diff_budget:,.0f} 円", delta_color="normal")
            
            if consumption_rate > 100:
                b_col4.metric("予算消化率", f"{consumption_rate:.1f} %", delta="⚠️ 予算超過！", delta_color="inverse")
            else:
                b_col4.metric("予算消化率", f"{consumption_rate:.1f} %", delta="正常範囲", delta_color="normal")

            df_vs = pd.DataFrame([
                {"カテゴリ": "内装(内)", "種類": "実行予算", "金額": b_a},
                {"カテゴリ": "内装(内)", "種類": "支払実績", "金額": p_a},
                {"カテゴリ": "設備(設)", "種類": "実行予算", "金額": b_b},
                {"カテゴリ": "設備(設)", "種類": "支払実績", "金額": p_b},
                {"カテゴリ": "電気(電)", "種類": "実行予算", "金額": b_c},
                {"カテゴリ": "電気(電)", "種類": "支払実績", "金額": p_c},
                {"カテゴリ": "厨房(P)", "種類": "実行予算", "金額": b_d},
                {"カテゴリ": "厨房(P)", "種類": "支払実績", "金額": p_d},
                {"カテゴリ": "依頼業社", "種類": "実行予算", "金額": b_req},
                {"カテゴリ": "依頼業社", "種類": "支払実績", "金額": p_req},
            ])

            chart_vs = alt.Chart(df_vs).mark_bar().encode(
                x=alt.X('カテゴリ:N', sort=["内装(内)", "設備(設)", "電気(電)", "厨房(P)", "依頼業社"], title=None),
                y=alt.Y('金額:Q', title="金額 (円)"),
                color=alt.Color(
                    '種類:N', 
                    scale=alt.Scale(domain=['実行予算', '支払実績'], range=['#3182ce', '#e53e3e']),
                    legend=alt.Legend(title="種別")
                ),
                xOffset='種類:N',
                tooltip=['カテゴリ', '種類', '金額']
            ).properties(height=250)

            st.altair_chart(chart_vs, use_container_width=True)

        bottom_left, bottom_right = st.columns([1, 1.8])
        
        with bottom_left:
            with st.container(border=True):
                st.markdown("#### 営業計算書メモ")
                in_eval = st.text_area("＜現場での評価＞", value=cur_data.get("eval_memo", ""), key=f"eval_{selected_project}_v24")
                in_budg = st.text_area("＜予算経経緯＞", value=cur_data.get("budget_memo", ""), height=150, key=f"budg_{selected_project}_v24")
                in_materials = st.text_area("＜建材・仕様の記録メモ＞", value=cur_data.get("materials_memo", ""), key=f"mat_{selected_project}_v24")

        with bottom_right:
            with st.container(border=True):
                st.markdown("#### 工事カテゴリ別内訳")
                s_Req, s_A, s_B, s_C, s_D = st.tabs(["🆕 依頼業社", "内 (内装)", "設 (設備)", "電 (電気)", "P (厨房)"])
                
                contractor_names = [str(n).strip() for n in df_contractor_master["業者・工種名"].tolist() if str(n).strip() != "" and str(n).strip().lower() != "nan"]
                used_names = []
                for cat in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"]:
                    for item in cur_data.get(cat, []):
                        n = str(item.get("業者・工種名", "")).strip()
                        if n and n not in contractor_names and n not in used_names: used_names.append(n)
                
                contractor_options = [""] + contractor_names + used_names
                
                # 💡改善ポイント: column_configで幅を指定し、Tabキーによる操作性を向上
                g_conf = {
                    "工種名": st.column_config.TextColumn("読取り種類", disabled=False, width="medium"), 
                    "業者管理番号": None,  
                    "業者・工種名": st.column_config.SelectboxColumn("業者・工種名", options=contractor_options, width="large"),
                    "実行予算": st.column_config.NumberColumn("実行予算", width="medium"),
                    "協力業者支払": st.column_config.NumberColumn("協力業者支払", width="medium"),
                    "完了金額": None
                }
                
                if "update_key" not in st.session_state: st.session_state.update_key = 0
                u_key = st.session_state.update_key
                
                def clean_req_df(data_list):
                    df = pd.DataFrame(data_list)
                    if df.empty: return pd.DataFrame(columns=["移動先", "工種名", "業者・工種名", "実行予算", "協力業者支払", "完了金額"])
                    df = df.drop(columns=["業者管理番号", "顧客管理番号"], errors='ignore')
                    df = df.fillna({"移動先": "-", "工種名": "", "業者・工種名": "", "実行予算": 0, "協力業者支払": 0, "完了金額": 0})
                    for col in ["実行予算", "協力業者支払", "完了金額"]:
                        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    return df
                    
                def clean_df(data_list):
                    df = pd.DataFrame(data_list)
                    if df.empty: return pd.DataFrame(columns=["工種名", "業者・工種名", "実行予算", "協力業者支払", "完了金額"])
                    df = df.drop(columns=["業者管理番号", "顧客管理番号"], errors='ignore')
                    df = df.fillna({"工種名": "", "業者・工種名": "", "実行予算": 0, "協力業者支払": 0, "完了金額": 0})
                    for col in ["実行予算", "協力業者支払", "完了金額"]:
                        if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    return df

                with s_Req:
                    req_conf = {
                        "移動先": st.column_config.SelectboxColumn("移動先", options=["-", "内", "設", "電", "P"], default="-", width="small"),
                        "工種名": st.column_config.TextColumn("読取り種類", width="medium"),
                        "業者・工種名": st.column_config.SelectboxColumn("業者・工種名", options=contractor_options, width="large"),
                        "実行予算": st.column_config.NumberColumn("実行予算", width="medium"),
                        "協力業者支払": st.column_config.NumberColumn("協力業者支払", width="medium"),
                        "完了金額": None
                    }
                    df_req = clean_req_df(cur_data.get("detail_request", []))
                    ed_req = st.data_editor(df_req, column_config=req_conf, num_rows="dynamic", key=f"ed_req_{selected_project}_{u_key}_v24", use_container_width=True, hide_index=True)
                    
                with s_A: ed_a = st.data_editor(clean_df(cur_data["detail_a"]), use_container_width=True, hide_index=True, column_config=g_conf, key=f"ea_{selected_project}_{u_key}_v24", num_rows="dynamic")
                with s_B: ed_b = st.data_editor(clean_df(cur_data["detail_b"]), use_container_width=True, hide_index=True, column_config=g_conf, key=f"eb_{selected_project}_{u_key}_v24", num_rows="dynamic")
                with s_C: ed_c = st.data_editor(clean_df(cur_data["detail_c"]), use_container_width=True, hide_index=True, column_config=g_conf, key=f"ec_{selected_project}_{u_key}_v24", num_rows="dynamic")
                with s_D: ed_d = st.data_editor(clean_df(cur_data["detail_d"]), use_container_width=True, hide_index=True, column_config=g_conf, key=f"ed_{selected_project}_{u_key}_v24", num_rows="dynamic")

        st.markdown("---")
        st.subheader("営業計算書（画像プレビュー）")
        calc_img_col1, calc_img_col2 = st.columns([1, 1])
        
        with calc_img_col1:
            saved_img = cur_data.get("calc_image", "")
            if saved_img and os.path.exists(saved_img): st.image(saved_img, use_container_width=True)
            else: st.info("現在、この物件に登録されている営業計算書の画像はありません。")
                
        with calc_img_col2: new_img = st.file_uploader("新しい営業計算書をアップロード", type=["png", "jpg", "jpeg", "pdf"], key=f"img_up_{selected_project}_v24")

        st.markdown("---")
        with st.expander("現場投入工数（就業時間）のメンバーを選択・表示する", expanded=False):
            col_st1, col_st2, col_st3, col_st4 = st.columns(4)
            with col_st1:
                st.markdown("**第1管理部（設備）- 設**")
                in_staff_setsubi = st.multiselect("担当メンバー選択", all_staff_list, default=cur_data.get("staff_setsubi", []), key=f"sel_setsubi_{selected_project}_v24")
                for s in in_staff_setsubi:
                    h = get_hours_from_andpad(s, selected_project, in_proj_num)
                    st.write(f"・{s}: `{h:.1f} h`")
            with col_st2:
                st.markdown("**第2管理部（内装）- 内**")
                in_staff_naisou = st.multiselect("担当メンバー選択", all_staff_list, default=cur_data.get("staff_naisou", []), key=f"sel_naisou_{selected_project}_v24")
                for s in in_staff_naisou:
                    h = get_hours_from_andpad(s, selected_project, in_proj_num)
                    st.write(f"・{s}: `{h:.1f} h`")
            with col_st3:
                st.markdown("**第3管理部（電気）- 電**")
                in_staff_denki = st.multiselect("担当メンバー選択", all_staff_list, default=cur_data.get("staff_denki", []), key=f"sel_denki_{selected_project}_v24")
                for s in in_staff_denki:
                    h = get_hours_from_andpad(s, selected_project, in_proj_num)
                    st.write(f"・{s}: `{h:.1f} h`")
            with col_st4:
                st.markdown("**PM室（厨房等）- P**")
                in_staff_pm = st.multiselect("担当メンバー選択", all_staff_list, default=cur_data.get("staff_pm", []), key=f"sel_pm_{selected_project}_v24")
                for s in in_staff_pm:
                    h = get_hours_from_andpad(s, selected_project, in_proj_num)
                    st.write(f"・{s}: `{h:.1f} h`")

        # 💡改善ポイント: 画面下部にも共通の保存処理ボタンを設置
        st.markdown("---")
        if st.button(f"💾 【{selected_project}】 を保存", use_container_width=True, type="primary", key=f"save_bottom_{selected_project}_v24") or st.session_state.get(f"save_top_{selected_project}", False):
            projects_db[selected_project]["client_name"] = in_client
            projects_db[selected_project]["project_number"] = in_proj_num; projects_db[selected_project]["project_year"] = in_year
            projects_db[selected_project]["project_month"] = in_month; 
            projects_db[selected_project]["bunrui_1"] = in_bunrui1
            projects_db[selected_project]["bunrui_2"] = in_bunrui2
            projects_db[selected_project]["motouke_shitauke"] = in_motouke
            projects_db[selected_project]["受注金額"] = in_order_amount
            projects_db[selected_project]["project_dept"] = in_project_dept
            projects_db[selected_project]["eval_memo"] = in_eval; projects_db[selected_project]["budget_memo"] = in_budg
            projects_db[selected_project]["materials_memo"] = in_materials
            
            def map_contractor_id(records):
                for r in records:
                    c_name = r.get("業者・工種名", "")
                    if c_name:
                        m = df_contractor_master[df_contractor_master["業者・工種名"] == c_name]
                        if not m.empty: r["業者管理番号"] = m.iloc[0]["業者管理番号"]
                        else: r["業者管理番号"] = ""
                    else: r["業者管理番号"] = ""
                return records

            projects_db[selected_project]["detail_a"] = map_contractor_id(ed_a.to_dict('records'))
            projects_db[selected_project]["detail_b"] = map_contractor_id(ed_b.to_dict('records'))
            projects_db[selected_project]["detail_c"] = map_contractor_id(ed_c.to_dict('records'))
            projects_db[selected_project]["detail_d"] = map_contractor_id(ed_d.to_dict('records'))
            
            ed_req_mapped = map_contractor_id(ed_req.to_dict('records'))
            new_req = []
            for row in ed_req_mapped:
                dest = row.get("移動先", "-")
                if not row.get("工種名") and not row.get("業者管理番号") and not row.get("業者・工種名"): continue
                if dest == "-": new_req.append(row)
                else:
                    row["移動先"] = "-"
                    if dest == "内": projects_db[selected_project]["detail_a"].append(row)
                    elif dest == "設": projects_db[selected_project]["detail_b"].append(row)
                    elif dest == "電": projects_db[selected_project]["detail_c"].append(row)
                    elif dest == "P": projects_db[selected_project]["detail_d"].append(row)
            
            projects_db[selected_project]["detail_request"] = new_req
            projects_db[selected_project]["staff_setsubi"] = in_staff_setsubi
            projects_db[selected_project]["staff_naisou"] = in_staff_naisou
            projects_db[selected_project]["staff_denki"] = in_staff_denki
            projects_db[selected_project]["staff_pm"] = in_staff_pm

            if new_img is not None:
                os.makedirs("uploaded_images", exist_ok=True)
                safe_name = "".join([c for c in selected_project if c.isalnum() or c in " _-"])
                img_path = f"uploaded_images/{safe_name}_calc.png"
                if new_img.name.lower().endswith(".pdf"):
                    try:
                        import fitz
                        doc = fitz.open(stream=new_img.read(), filetype="pdf")
                        page = doc.load_page(0)
                        pix = page.get_pixmap(dpi=150)
                        img_to_save = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        img_to_save.save(img_path)
                    except:
                        pass
                else:
                    with open(img_path, "wb") as f: f.write(new_img.getbuffer())
                projects_db[selected_project]["calc_image"] = img_path

            save_projects(projects_db)
            st.success("保存完了！選択された担当系統と受注金額がロックされ、全体の売上集計へ反映されました。")
            st.session_state.update_key += 1
            st.rerun()

# ==========================================
# --- タブ3：AI書類自動処理 ＆ ANDPADデータ取込 ---
# ==========================================
with tab3:
    st.header("AI書類自動処理 ＆ 担当表取込")
    st.markdown("書類やExcelファイルをアップロードして、システムに自動で数値を読み込ませます。")

    if "ai_result" not in st.session_state: st.session_state.ai_result = None

    # 💡改善ポイント: UIをステップ形式（枠線）で展開し、作業に迷わないようにする
    with st.container(border=True):
        st.markdown("### Step 1: 書類の選択とアップロード")
        col_doc, col_input = st.columns(2)
        with col_doc: 
            doc_type = st.radio("📄 読み込む種類", ["営業計算書（新規物件登録）", "発注書", "請求書等", "担当表(ANDPAD Excel)"], key="ai_doc_type_v24")
        with col_input: 
            if doc_type == "担当表(ANDPAD Excel)":
                st.info("Excelデータは「ファイルから選ぶ」のみ対応しています。")
                input_type = "ファイルから選ぶ"
            else:
                input_type = st.radio("📷 入力方法", ["ファイル(PDF/画像)から選ぶ", "カメラで撮影する"], key="ai_input_type_v24")

        # ANDPADの処理
        if doc_type == "担当表(ANDPAD Excel)":
            excel_file = st.file_uploader("担当表Excel (.xlsx, .xls) を選択", type=["xlsx", "xls"], key="ai_excel_uploader_v24")
            if excel_file is not None:
                if st.button("🚀 担当表をシステムに反映", type="primary", use_container_width=True):
                    os.makedirs("uploaded_images", exist_ok=True)
                    current_ymd = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    acc_path = os.path.join(ANDPAD_ACC_DIR, f"andpad_{current_ymd}.xlsx")
                    with open(acc_path, "wb") as f: f.write(excel_file.getbuffer())
                    get_accumulated_andpad_dict.clear()
                    b_path = backup_uploaded_file(excel_file.getvalue(), excel_file.name, prefix="attendance")
                    st.success(f"✅ 担当表データの蓄積とバックアップが完了しました！")
                    st.rerun()

            if os.path.exists(ANDPAD_ACC_DIR) and os.listdir(ANDPAD_ACC_DIR):
                st.markdown("---")
                if st.button("🗑️ 蓄積されているANDPADデータをリセット", type="secondary", key="reset_andpad_acc_btn"):
                    for f in os.listdir(ANDPAD_ACC_DIR):
                        try: os.remove(os.path.join(ANDPAD_ACC_DIR, f))
                        except: pass
                    get_accumulated_andpad_dict.clear()
                    st.success("🧹 蓄積ANDPADデータをリセットしました！")
                    st.rerun()

                extracted_projs_with_hours = extract_projects_and_hours_from_dict()
                if extracted_projs_with_hours:
                    st.markdown("#### 現在システムが認識しているExcel内の案件名と【抽出合計時間】")
                    st.code(" \n".join(extracted_projs_with_hours))

        # 通常書類の処理
        else:
            image_file = None
            if input_type == "ファイル(PDF/画像)から選ぶ": 
                image_file = st.file_uploader("画像（jpg, png）または PDF", type=["jpg", "jpeg", "png", "pdf"], key="ai_file_uploader_v24")
            else: 
                image_file = st.camera_input("書類を撮影", key="ai_camera_input_v24")

    # 💡 Step 2: 画像がアップロードされたらAI処理画面を開く
    if doc_type != "担当表(ANDPAD Excel)" and image_file is not None:
        with st.container(border=True):
            st.markdown("### Step 2: AI解析の実行")
            if input_type == "ファイル(PDF/画像)から選ぶ" and image_file.name.lower().endswith('.pdf'):
                try:
                    import fitz
                    doc = fitz.open(stream=image_file.read(), filetype="pdf")
                    page = doc.load_page(0)
                    pix = page.get_pixmap(dpi=180)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                except ImportError:
                    st.error("🚨 PDF読み込み機能を使うためのライブラリが不足しています。\nVS Codeのターミナルで `pip install PyMuPDF` と入力して実行し、アプリを再起動してください。")
                    st.stop()
            else:
                img = Image.open(image_file).convert('RGB')
            
            img.thumbnail((2048, 2048), Image.Resampling.LANCZOS)
            enhancer = ImageEnhance.Contrast(img)
            img_processed = enhancer.enhance(1.2)
            sharpener = ImageEnhance.Sharpness(img_processed)
            img_processed = sharpener.enhance(1.3)
                
            st.image(img_processed, caption="AIが解析する調整済み画像", width=400)
            
            if st.button("✨ AIで情報を自動抽出する", type="primary", use_container_width=True, key="ai_extract_btn_v24"):
                if GEMINI_API_KEY == "ここに取得したAPIキーを貼り付けてください" or GEMINI_API_KEY == "YOUR_KEY_HERE":
                    st.error("⚠️ コードの上部にある `GEMINI_API_KEY` が設定されていません。先ほど取得したキーを貼り付けてください。")
                else:
                    with st.spinner('AIが高精度解析中...（十数秒かかる場合があります）'):
                        try:
                            # 💡【最終改善】2段構成の不規則性（クリーンサトウ問題）を完全に突破するFew-Shotプロンプト
                            sys_instruction = """
                            あなたは建築・内装業界の精密な帳票・営業計算書解析に特化した最高精度の経理AIアシスタントです。
                            画像内の文字、表、数字、手書きの注記を極めて高い精度で正確に解析し、指定されたJSON構造のみを出力してください。
                            
                            【🚨 悪質な「行ズレ」を完全に防ぐための最終ルール】
                            書類の表において、1つの発注項目が「工種」と「業者名」の2段書きになっていますが、**業者名が上段に来るか下段に来るかが非常に不規則**です。
                            これによってAIが行ズレを起こし、下の業者の金額を上にコピーしてしまう深刻なエラーが発生しています。
                            
                            ＜実際の不規則な例＞
                            パターンA（業者が下段）：
                            「クリーニング」
                            「クリーンサトウ」 → 業者名は下段だが、金額は 80,000
                            
                            パターンB（業者が上段）：
                            「佐々木内装」
                            「金物」 → 業者名は上段で、金額は 800,000
                            
                            【厳守事項】
                            1. 業者名が上段・下段のどちらにあっても、必ず「その文字が含まれるひとまとまりのブロック（同じ行・枠内）」にある金額を正確に抽出してください。
                            2. 隣接する上下のブロックの金額を絶対に混同しないでください（例：「クリーンサトウ」に800,000を割り当てるのは重大なエラーです）。
                            3. 全ての業者を抽出した後、金額が上下の業者と重複していないか（行ズレしていないか）必ず自己チェックを行ってから出力してください。
                            4. 「小計」「合計」の行は絶対に抽出しないでください。
                            5. 金額は必ず「税抜」で、カンマなしの半角数字のみを出力してください。
                            """
                            
                            model = genai.GenerativeModel('gemini-2.5-flash', system_instruction=sys_instruction)
                            
                            prompt_base = f"""
                            添付された画像は「{doc_type}」です。画像から必要な項目を完璧に抽出して、JSON形式で出力してください。
                            
                            【重要抽出ルール】
                            1. 手書き文字、枠線内の数字、薄い印刷文字も取りこぼさず読み取ってください。
                            2. 表の行を横にスキャンし、**「業者名」と「金額」の組み合わせ（行）が絶対にズレないように**してください。上下2段の記載に騙されないでください。
                            3. ⚠️「発注先」が「小計」や「合計」となっている行は絶対に無視してください。
                            4. 記載されているすべての業者名、材料費、外注費、現場経費の行を漏らさず全てリスト化してください。
                            """
                            
                            if doc_type == "営業計算書（新規物件登録）":
                                # 💡【分類・元請/下請のルールを強力に指定】
                                prompt = prompt_base + """
                                次のフォーマットのJSONのみを出力してください：
                                {
                                    "物件番号": "抽出した物件番号(例: 56T-2-3005など)",
                                    "顧客名": "抽出した契約先の会社名",
                                    "物件名": "抽出した現場名",
                                    "受注金額": 1234567, 
                                    "月度": "抽出した月度（例：7月など。上部の『受注・請求(出来高請求)』の行で金額が入っている列の月）",
                                    "分類1": "抽出した分類1（手書きの丸印がついている項目。例：C.外食産業）",
                                    "分類2": "抽出した分類2（手書きの丸印がついている項目。例：C.内装工事）",
                                    "元請_下請": "抽出した元請・下請の情報（手書きの丸印がついている項目。例：元請）",
                                    "材料費リスト": [ {"会社名": "〇〇建材", "金額": 100000} ],
                                    "外注費リスト": [ {"会社名": "〇〇設備", "金額": 250000} ],
                                    "現場経費リスト": [ {"会社名": "〇〇リース", "金額": 15000} ]
                                }
                                
                                🎯【「受注金額」の絶対抽出ルール】
                                - ⚠️【超重要】必ず「税抜（消費税を含まない）」の「本工事受注金額」を抽出してください。
                                - 「税込価格」や「受注合計」として記載されている金額（税率が足された後の大きい金額）は絶対に抽出しないでください。
                                
                                🎯【「分類・元請/下請」の絶対抽出ルール】
                                - 複数の選択肢が並んでいる項目（分類1、分類2、元請/下請など）は、**必ず「手書きの丸（〇）や楕円の線」で囲まれている選択肢**を抽出してください。丸がついていない文字を推測で選んではいけません。
                                
                                🎯【「業者・金額リスト」の絶対抽出ルール】
                                - 表の「形態/工種」「発注先」列と、「実行予算組」「予算残高」列を照らし合わせてください。
                                - 各業者の行を横にスキャンし、**必ず同じ行にある金額**を正確に抽出してください。
                                - 全ての企業名・項目名と金額を1行残らず完全網羅して抽出してください。
                                """
                            elif doc_type == "発注書":
                                prompt = prompt_base + """
                                {
                                   "物件番号": "抽出した物件番号",
                                    "物件名": "抽出した物件名",
                                    "業社・工種名": "三岳工業株式会社以外の会社名",
                                    "顧客名": "抽出した顧客名",
                                    "発注金額": 123456
                                }
                                ※金額は「税抜振込金額」の数字のみを数値(カンマなし)で入れてください。
                                """
                            else:
                                prompt = prompt_base + """
                                {
                                    "物件名": "抽出した物件名",
                                    "業者名": "抽出した業者名",
                                    "金額": 123456
                                }
                                ※金額は「税抜支払金額」の数字のみを数値(カンマなし)で入れてください。
                                """
                            
                            response = model.generate_content(
                                [prompt, img_processed],
                                generation_config={"response_mime_type": "application/json", "temperature": 0.0}
                            )
                            res_text = response.text
                            
                            try:
                                extracted_data = json.loads(res_text)
                                st.session_state.ai_result = extracted_data
                                st.success("🎉 高精度AIによる読み取りが完了しました！下部のStep3に進んでください。")
                            except json.JSONDecodeError:
                                json_str = re.search(r'\{.*\}', res_text, re.DOTALL)
                                if json_str:
                                    extracted_data = json.loads(json_str.group())
                                    st.session_state.ai_result = extracted_data
                                    st.success("🎉 高精度AIによる読み取りが完了しました！下部のStep3に進んでください。")
                                else:
                                    st.error("読み取りフォーマットの解析に失敗しました。再度お試しください。")
                            
                        except Exception as e: st.error(f"エラーが発生しました: {e}")
     
    # 💡 Step 3: AI結果が存在する時のみ確認と反映のUIを表示する
    if st.session_state.ai_result:
        with st.container(border=True):
            st.markdown("### Step 3: 抽出結果の確認と反映")
            data = st.session_state.ai_result
            st.write("AIが読み取ったデータ:")
            st.json(data)
            st.info("👇 以下の項目を確認・修正して反映ボタンを押してください。")
            
            project_names = list(projects_db.keys())
            
            if doc_type == "営業計算書（新規物件登録）":
                confirm_proj = st.text_input("新規物件名", value=data.get("物件名", ""), key="ai_proj_name_v24")
                confirm_client = st.text_input("顧客名", value=data.get("顧客名", ""), key="ai_client_name_v24")
                confirm_num = st.text_input("物件番号", value=data.get("物件番号", ""), key="ai_proj_num_v24")
                
                guessed_term = "57期"
                if confirm_num and confirm_num[:2].isdigit():
                    term_str = confirm_num[:2] + "期"
                    if term_str in YEAR_LIST: guessed_term = term_str
                        
                confirm_year = st.selectbox("対象の期", YEAR_LIST, index=YEAR_LIST.index(guessed_term), key="ai_proj_year_v24")
                
                ai_month = data.get("月度", "9月")
                if ai_month not in MONTH_LIST: ai_month = "9月"
                confirm_month = st.selectbox("月度（売上計上月）", MONTH_LIST, index=MONTH_LIST.index(ai_month), key="ai_proj_month_v24")
                
                col_b1, col_b2, col_b3 = st.columns(3)
                with col_b1:
                    confirm_bunrui1 = st.text_input("分類1", value=data.get("分類1", ""), key="ai_proj_bunrui1_v24")
                with col_b2:
                    confirm_bunrui2 = st.text_input("分類2", value=data.get("分類2", ""), key="ai_proj_bunrui2_v24")
                with col_b3:
                    confirm_motouke = st.text_input("元請/下請", value=data.get("元請_下請", ""), key="ai_proj_motouke_v24")
                
                st.markdown("#### 📋 抽出された業者・金額リスト")
                st.info("AIが読み取った材料費・外注費・現場経費のリストです。内容の修正、行の追加・削除が可能です。")
                
                extracted_items = []
                for mat_key in ["材料費リスト", "材料費"]:
                    for mat in data.get(mat_key, []):
                        name = mat.get("会社名", mat.get("業者名", mat.get("名称", mat.get("項目", ""))))
                        amt = mat.get("金額", mat.get("価格", mat.get("費用", 0)))
                        if name and str(name).lower() != "none" and str(name).strip() != "":
                            extracted_items.append({"工種名": "AI読取(材料費)", "業者・工種名": str(name).strip(), "金額": safe_num(amt)})
                
                for sub_key in ["外注費リスト", "外注費"]:
                    for sub in data.get(sub_key, []):
                        name = sub.get("会社名", sub.get("業者名", sub.get("名称", sub.get("項目", ""))))
                        amt = sub.get("金額", sub.get("価格", sub.get("費用", 0)))
                        if name and str(name).lower() != "none" and str(name).strip() != "":
                            extracted_items.append({"工種名": "AI読取(外注費)", "業者・工種名": str(name).strip(), "金額": safe_num(amt)})

                for exp_key in ["現場経費リスト", "現場経費"]:
                    for exp in data.get(exp_key, []):
                        name = exp.get("会社名", exp.get("業者名", exp.get("名称", exp.get("項目", ""))))
                        amt = exp.get("金額", exp.get("価格", exp.get("費用", 0)))
                        if name and str(name).lower() != "none" and str(name).strip() != "":
                            extracted_items.append({"工種名": "AI読取(現場経費)", "業者・工種名": str(name).strip(), "金額": safe_num(amt)})
                
                df_extracted = pd.DataFrame(extracted_items) if extracted_items else pd.DataFrame(columns=["工種名", "業者・工種名", "金額"])
                edited_df = st.data_editor(df_extracted, num_rows="dynamic", use_container_width=True, key="ai_multi_items_ed")
                
                if st.button("🚀 この営業計算書データを反映する（新規作成 / 既存上書き）", type="primary", key="ai_btn_v24"):
                    if confirm_proj:
                        is_new = confirm_proj not in projects_db
                        if is_new:
                            p_data = get_default_project_data()
                            p_data["budget_memo"] = ""
                        else:
                            p_data = projects_db[confirm_proj]
     
                        if confirm_client: p_data["client_name"] = confirm_client
                        if confirm_num: p_data["project_number"] = confirm_num
                        p_data["project_year"] = confirm_year
                        p_data["project_month"] = confirm_month
                        
                        p_data["bunrui_1"] = confirm_bunrui1
                        p_data["bunrui_2"] = confirm_bunrui2
                        p_data["motouke_shitauke"] = confirm_motouke
                        
                        try: order_amt = int(str(data.get("受注金額", 0)).replace(',', ''))
                        except: order_amt = 0
                        p_data["受注金額"] = order_amt
                        
                        update_msg = f"【{datetime.datetime.now().strftime('%Y/%m/%d %H:%M')} AI読取】本工事受注金額: {order_amt:,} 円 (月度: {confirm_month})"
                        if is_new:
                            p_data["budget_memo"] = update_msg
                        else:
                            p_data["budget_memo"] = update_msg + "\n\n" + str(p_data.get("budget_memo", ""))
                        
                        new_extracted_rows = []
                        for _, row in edited_df.iterrows():
                            g_name = str(row.get("業者・工種名", "")).strip()
                            amt = safe_num(row.get("金額", 0))
                            k_name = str(row.get("工種名", "AI読取(その他)")).strip()
                            if g_name or amt > 0:
                                new_extracted_rows.append({
                                    "移動先": "-", "工種名": k_name, "業者管理番号": "", "業者・工種名": g_name,
                                    "実行予算": 0, "協力業者支払": amt, "完了金額": 0
                                })
                        
                        if new_extracted_rows:
                            existing_req = p_data.get("detail_request", [])
                            non_ai_req = [req for req in existing_req if not str(req.get("工種名", "")).startswith("AI読取")]
                            p_data["detail_request"] = non_ai_req + new_extracted_rows
                        
                        projects_db[confirm_proj] = p_data
                        save_projects(projects_db)
                        
                        if is_new:
                            st.success(f"🎉 新規物件「{confirm_proj}」を登録しました！")
                        else:
                            st.success(f"🔄 既存物件「{confirm_proj}」のデータを最新版に上書き更新しました！")
                            
                        st.session_state.ai_result = None # UIリセット
                        st.rerun()
                    else:
                        st.error("物件名を入力してください。")
                        
            elif doc_type == "発注書":
                selected_proj = st.selectbox("紐付ける物件名", ["(選択してください)"] + project_names, key="ai_sel_proj_発注_v24")
                confirm_client = st.text_input("顧客名", value=data.get("顧客名", ""), key="ai_client_発注_v24")
                
                ai_amount = data.get("発注金額", data.get("金額", 0))
                confirm_amount = st.number_input("税抜振込金額", value=int(ai_amount), key="ai_amt_発注_v24")
                ai_contractor = data.get("業社・工種名", "")
                confirm_contractor = st.text_input("業者名", value=ai_contractor, key="ai_contractor_発注_v24")
                
                if st.button("🚀 この発注データを物件に反映する", type="primary", key="ai_btn_発注_v24"):
                    if selected_proj != "(選択してください)":
                        projects_db[selected_proj]["client_name"] = confirm_client
                        
                        new_req = {
                            "移動先": "-", 
                            "工種名": "AI読取(発注書)", 
                            "業者管理番号": "", 
                            "業者・工種名": confirm_contractor,
                            "実行予算": 0, 
                            "協力業者支払": confirm_amount, 
                            "完了金額": 0
                        }
                        if "detail_request" not in projects_db[selected_proj]:
                            projects_db[selected_proj]["detail_request"] = []
                        projects_db[selected_proj]["detail_request"].append(new_req)
                        
                        save_projects(projects_db)
                        st.success(f"✅ {selected_proj} に発注データ（{confirm_contractor}：{confirm_amount}円）を反映しました！")
                        st.session_state.ai_result = None
                        st.rerun()
                    else: st.error("物件名を選択してください。")
                    
            else:
                selected_proj = st.selectbox("紐付ける物件名", ["(選択してください)"] + project_names, key="ai_sel_proj_請求_v24")
                ai_contractor = data.get("業者名", data.get("業社・工種名", ""))
     
                c_col1, c_col2, c_col3 = st.columns(3)
                with c_col1:
                    confirm_gyousha = st.text_input("紐付ける業者", value=ai_contractor, key="ai_sel_gyo_請求_v24")
                with c_col2:
                    ai_koshu = data.get("工種名", "AI自動読取")
                    confirm_koshu = st.text_input("工種名（メモ）", value=ai_koshu, key="ai_koshu_請求_v24")
                with c_col3:
                    confirm_amount = st.number_input("税抜支払金額", value=int(data.get("金額", data.get("発注金額", 0))), key="ai_amt_請求_v24")
                
                if st.button("🚀 この支払データを「依頼業社」に送る", type="primary", key="ai_btn_請求_v24"):
                    if selected_proj != "(選択してください)":
                        new_req = {
                            "移動先": "-", "工種名": confirm_koshu, "業者管理番号": "", "業者・工種名": confirm_gyousha,
                            "実行予算": 0, "協力業者支払": confirm_amount, "完了金額": 0
                        }
                        
                        projects_db[selected_project]["detail_request"].append(new_req)
                        save_projects(projects_db)
                        st.success(f"🎉 大成功！ {selected_proj} の「依頼業社」タブにデータを送信しました！")
                        st.session_state.ai_result = None
                        st.rerun()
                    else: st.error("物件名を選択してください。")
 
# ==========================================
# --- タブ4：顧客別 実績・分析 ---
# ==========================================
with tab4:
    st.header("📈 主要顧客別 実績・分析")
    st.markdown("日々の物件データからの自動集計と、過去データのエクセル取込を行います。")
    
    t4_sub1, t4_sub2 = st.tabs(["📊 現在・未来の自動集計 (57期〜)", "📁 過去実績データの取込・自動振り分け (〜56期)"])
    
    with t4_sub1:
        st.subheader("現在入力されているデータからの顧客別 自動集計")
        
        all_clients = sorted(list(set([v.get("client_name", "未設定") for v in projects_db.values() if isinstance(v, dict)])))
        all_clients = [c for c in all_clients if c != "未設定" and "合計" not in c and "修理" not in c and "スキップ" not in c]
        
        col_c1, col_c2 = st.columns(2)
        with col_c1: sel_client = st.selectbox("分析する顧客名を選択", ["(選択してください)"] + all_clients, key="client_anal_sel_v24")
        with col_c2: sel_year = st.selectbox("対象の期を選択（自動集計用）", YEAR_LIST, index=YEAR_LIST.index("57期") if "57期" in YEAR_LIST else 0, key="year_anal_sel_v24")
            
        if sel_client != "(選択してください)":
            client_sales = 0.0
            client_cost = 0.0
            month_sales = {m: 0.0 for m in MONTH_LIST}
            month_cost = {m: 0.0 for m in MONTH_LIST}
            target_projects = []
            
            for p_name, p_data in projects_db.items():
                if not isinstance(p_data, dict): continue
                if p_data.get("client_name") == sel_client and p_data.get("project_year") == sel_year:
                    target_projects.append(p_name)
                    p_month = p_data.get("project_month", "9月")
                    
                    val_a = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_a", []))
                    val_b = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_b", []))
                    val_c = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_c", []))
                    val_d = sum(safe_num(x.get("完了金額")) for x in p_data.get("detail_d", []))
                    
                    order_amt = safe_num(p_data.get("受注金額", 0))
                    p_sales = order_amt if order_amt > 0 else (val_a + val_b + val_c + val_d)
                    
                    p_cost = 0.0
                    for cat in ["detail_a", "detail_b", "detail_c", "detail_d", "detail_request"]:
                        for item in p_data.get(cat, []): p_cost += safe_num(item.get("協力業者支払"))
                            
                    client_sales += p_sales
                    client_cost += p_cost
                    if p_month in month_sales:
                        month_sales[p_month] += p_sales
                        month_cost[p_month] += p_cost
                        
            client_profit = client_sales - client_cost
            client_margin = (client_profit / client_sales) if client_sales > 0 else 0.0
            
            st.markdown(f"### 🏢 {sel_client} - {sel_year} 実績")
            st.write(f"**対象物件数:** {len(target_projects)}件 ({', '.join(target_projects)})")
            
            st.markdown("#### 📅 期 総合計")
            cols = st.columns(3)
            cols[0].metric("純売上額", f"{client_sales:,.0f} 円")
            cols[1].metric("粗利額", f"{client_profit:,.0f} 円")
            cols[2].metric("粗利率", f"{client_margin*100:.1f} %")
            
            row_sales = {"項目": "売上"}
            row_profit = {"項目": "粗利"}
            row_margin = {"項目": "粗利率"}
            
            for m in MONTH_LIST:
                row_sales[m] = f"{month_sales[m]:,.0f}"
                m_profit = month_sales[m] - month_cost[m]
                row_profit[m] = f"{m_profit:,.0f}"
                m_margin = (m_profit / month_sales[m]) if month_sales[m] > 0 else 0.0
                row_margin[m] = f"{m_margin*100:.1f}%"
                
            df_client_months = pd.DataFrame([row_sales, row_profit, row_margin])
            st.markdown("#### 📆 月別推移（9月〜8月）")
            st.info("※「その他」など、月別データが存在せず年間合計のみの顧客は、代表して「9月」の列に1年分の数字がまとめて表示されます。")
            st.dataframe(df_client_months, use_container_width=True, hide_index=True)
 
    with t4_sub2:
        st.subheader("📁 過去のExcelデータのアップロード・自動振り分け")
        st.info("※もし間違ったデータを取り込んでしまった場合は、下のボタンで一度リセットできます。")
        if st.button("🗑️ 【やり直し用】取り込んだ過去のデータをすべてリセット", type="secondary", key="reset_past_btn_v24"):
            keys_to_delete = [k for k in projects_db.keys() if str(k).startswith("【過去実績】")]
            if keys_to_delete:
                for k in keys_to_delete:
                    del projects_db[k]
                    # 🔽 Firestore上のデータも削除
                    db.collection("projects").document(k).delete()
                st.success(f"🧹 {len(keys_to_delete)}件の過去データをリセットしました。もう一度下のボタンから取り込んでください！")
                st.rerun()
            else: st.info("削除する過去データはありませんでした。")
                
        st.markdown("---")
        past_file = st.file_uploader("過去実績用エクセルファイル (.xlsx) を選択", type=["xlsx"], key="past_excel_file_v24")
        PAST_EXCEL_PATH = "uploaded_images/past_sales_data.xlsx"
        
        if past_file is not None:
            os.makedirs("uploaded_images", exist_ok=True)
            with open(PAST_EXCEL_PATH, "wb") as f: f.write(past_file.getbuffer())
            st.success("✅ 過去の売上データをアプリ内に保存しました！")
            
        if os.path.exists(PAST_EXCEL_PATH):
            try:
                df_past = load_excel_data(PAST_EXCEL_PATH)
                st.markdown("**保存されている過去データ（プレビュー）**")
                df_past_display = df_past.fillna("").astype(str)
                st.dataframe(df_past_display, use_container_width=True, height=250)
                
                st.markdown("---")
                if st.button("🚀 このExcelから【過去実績】を全自動抽出して登録する", type="primary", key="past_extract_btn_v24"):
                    add_count = 0
                    current_client = "未設定"
                    
                    month_cols = {
                        "9月": (16, 17), "10月": (19, 20), "11月": (22, 23), "12月": (25, 26),
                        "1月": (28, 29), "2月": (31, 32), "3月": (34, 35), "4月": (37, 38),
                        "5月": (40, 41), "6月": (43, 44), "7月": (46, 47), "8月": (49, 50)
                    }
                    
                    for i in range(len(df_past)):
                        col_A_val = str(df_past.iloc[i, 0]).strip()
                        if col_A_val != "nan" and col_A_val != "":
                            if "合計" in col_A_val or "修理" in col_A_val: current_client = "スキップ"
                            else: current_client = col_A_val
                                
                        col_B_val = str(df_past.iloc[i, 1]).strip()
                        match = re.search(r'(\d{2}期)', col_B_val)
                        
                        if match and current_client != "未設定" and current_client != "スキップ":
                            term_str = match.group(1)
                            data_row = i + 1
                            
                            if data_row >= len(df_past): continue
                            
                            month_data_found = False
                            
                            for m_name, (col_sales, col_profit) in month_cols.items():
                                if col_profit >= len(df_past.columns): continue
                                
                                try:
                                    sales_val = float(str(df_past.iloc[data_row, col_sales]).replace(',', '').strip())
                                    profit_val = float(str(df_past.iloc[data_row, col_profit]).replace(',', '').strip())
                                except: continue
                                    
                                if pd.isna(sales_val) or sales_val == 0: continue
                                    
                                month_data_found = True
                                cost_val = sales_val - profit_val
                                proj_name = f"【過去実績】{current_client}_{term_str}_{m_name}"
                                
                                if proj_name not in projects_db:
                                    new_item = get_default_project_data()
                                    new_item["client_name"] = current_client
                                    new_item["project_number"] = f"過去{term_str}"
                                    new_item["project_year"] = term_str
                                    new_item["project_month"] = m_name
                                    new_item["bunrui_1"] = "未設定"
                                    new_item["bunrui_2"] = "未設定"
                                    new_item["motouke_shitauke"] = "未設定"
                                    new_item["status"] = "アーカイブ（完了）"
                                    new_item["project_dept"] = "内"
                                    new_item["budget_memo"] = f"過去のExcelデータからの月別自動取り込み ({current_client})"
                                    
                                    for item in new_item["detail_a"]:
                                        if item["工種名"] == "内装工事":
                                            item["完了金額"] = sales_val
                                            item["協力業者支払"] = cost_val
                                            item["業者管理番号"] = "業001"
                                            break
                                    projects_db[proj_name] = new_item
                                    add_count += 1
                                    
                            if not month_data_found:
                                try:
                                    total_sales = float(str(df_past.iloc[i, 4]).replace(',', '').strip())
                                    total_profit = float(str(df_past.iloc[i, 5]).replace(',', '').strip())
                                except: continue
                                    
                                if not pd.isna(total_sales) and total_sales != 0:
                                    cost_val = total_sales - total_profit
                                    proj_name = f"【過去実績】{current_client}_{term_str}_通期"
                                    
                                    if proj_name not in projects_db:
                                        new_item = get_default_project_data()
                                        new_item["client_name"] = current_client
                                        new_item["project_number"] = f"過去{term_str}"
                                        new_item["project_year"] = term_str
                                        new_item["project_month"] = "9月"
                                        new_item["bunrui_1"] = "未設定"
                                        new_item["bunrui_2"] = "未設定"
                                        new_item["motouke_shitauke"] = "未設定"
                                        new_item["status"] = "アーカイブ（完了）"
                                        new_item["project_dept"] = "内"
                                        new_item["budget_memo"] = f"過去Excel（月別なし）からの年間合計取り込み ({current_client})"
                                        
                                        for item in new_item["detail_a"]:
                                            if item["工種名"] == "内装工事":
                                                item["完了金額"] = total_sales
                                                item["協力業者支払"] = cost_val
                                                item["業者管理番号"] = "業001"
                                                break
                                        projects_db[proj_name] = new_item
                                        add_count += 1
                                    
                    if add_count > 0:
                        save_projects(projects_db)
                        st.success(f"🎉 大成功！ {add_count}件の過去データをデータベースに登録しました！")
                        st.info("💡 「タブ1」に戻って、集計する期を変更して確認してください。")
                    else: st.info("新しく登録できるデータは見つかりませんでした。")
                        
            except Exception as e: st.error(f"Excelの読み込みに失敗しました: {e}")