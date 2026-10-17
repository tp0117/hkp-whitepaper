import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from google import genai
from google.genai import types
import json
import os
import time
import re
import io
import zipfile
import urllib.parse
from datetime import datetime
from PIL import Image

# 嘗試載入 python-docx 用於生成 Word 檔
try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except Exception:
    docx = None

# 1. 網頁標題及頁面設定
st.set_page_config(page_title="交通意外資料處理系統 (TAPI / TADO / T24)", page_icon="🚓", layout="wide")

# 2. UI 手機化美化 (Custom CSS)
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    div.stButton > button:first-child {
        width: 100%;
        height: 3.2rem;
        font-size: 1.2rem !important;
        font-weight: bold;
        background-color: #003366 !important;
        color: white !important;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .wa-button {
        display: block;
        width: 100%;
        height: 3.2rem;
        line-height: 3.2rem;
        text-align: center;
        background-color: #25D366;
        color: white !important;
        font-size: 1.1rem;
        font-weight: bold;
        border-radius: 10px;
        text-decoration: none;
        margin-top: 10px;
        margin-bottom: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    div[role="radiogroup"] {
        padding: 10px;
        background-color: #f0f2f6;
        border-radius: 8px;
    }
    div[role="radiogroup"] * {
        color: #0c1a30 !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# Session State 狀態初始化
state_keys = {
    "generated": False,
    "current_mode": "",
    "main_bytes": None,
    "main_filename": "",
    "main_mime": "",
    "tado_bytes": None,
    "tado_filename": "",
    "t24_2_bytes": None,
    "t24_2_filename": "",
    "zip_bytes": None,
    "zip_filename": "",
    "wa_url": "",
    "preview_data": None
}
for key, default_val in state_keys.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# Helper 函數：身份證號碼尾數加上圓括號
def format_hkid_parentheses(code_str):
    if not code_str:
        return ""
    s = str(code_str).strip()
    if re.search(r'\([0-9A-Za-z]\)$', s):
        return s
    match = re.match(r'^([A-Za-z]{1,2}\d{6})[\s\-_]?([0-9A-Za-z])$', s)
    if match:
        return f"{match.group(1)}({match.group(2)})"
    return s

# Helper 函數：日期解析與中文/星期轉換
def parse_date_and_weekday(date_str):
    if not date_str or date_str == "———":
        return "———", "———", "———"
    
    s = str(date_str).strip()
    dt = None
    
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%Y.%m.%d"]:
        try:
            dt = datetime.strptime(s, fmt)
            break
        except ValueError:
            pass
            
    if dt:
        weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
        weekday_zh = weekday_map[dt.weekday()]
        date_zh = f"{dt.year}年{dt.month}月{dt.day}日"
        date_std = dt.strftime("%d/%m/%Y")
        return date_std, date_zh, weekday_zh
    
    match_zh = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', s)
    if match_zh:
        y, m, d = int(match_zh.group(1)), int(match_zh.group(2)), int(match_zh.group(3))
        dt = datetime(y, m, d)
        weekday_map = {0: "一", 1: "二", 2: "三", 3: "四", 4: "五", 5: "六", 6: "日"}
        weekday_zh = weekday_map[dt.weekday()]
        return f"{d:02d}/{m:02d}/{y}", f"{y}年{m}月{d}日", weekday_zh

    return s, s, "———"

# Helper 函數：替換 Word 文件中的 Tag
def replace_docx_tags(doc, context):
    def replace_in_p(p):
        full_text = p.text
        for k, v in context.items():
            if k in full_text:
                full_text = full_text.replace(k, str(v) if (v is not None and str(v).strip() != "") else "———")
        if full_text != p.text:
            p.text = full_text

    for p in doc.paragraphs:
        replace_in_p(p)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    replace_in_p(p)

# 🔒 密碼鎖設定
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "8888")

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 白皮書 / TADO / T24 系統 - 安全驗證")
    st.info("本系統僅供內部授權人員使用，請輸入訪問密碼。")
    
    input_pwd = st.text_input("請輸入訪問密碼：", type="password")
    if st.button("🔑 登入"):
        if input_pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.success("驗證成功！正在載入系統...")
            st.rerun()
        else:
            st.error("密碼錯誤，請重新輸入！")
    st.stop()

# ----------------- 系統主介面 -----------------

st.title("🚓 交通意外資料處理系統 (TAPI / TADO / T24)")

top_col1, top_col2 = st.columns([3, 1])
with top_col2:
    if st.button("🔄 重置頁面"):
        for k in state_keys:
            st.session_state[k] = state_keys[k]
        st.rerun()

default_api_key = st.secrets.get("GEMINI_API_KEY", "")

with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input("Gemini API Key", value=default_api_key, type="password")
    st.markdown("[按此免費獲取 Google Gemini API Key](https://aistudio.google.com/)")
    
    st.divider()
    selected_model = st.selectbox(
        "🤖 首選 AI 模型",
        ["gemini-3.8-flash", "gemini-1.5-flash", "gemini-2.5-pro"],
        index=0
    )
    
    st.divider()
    if st.button("🔒 登出系統"):
        st.session_state["authenticated"] = False
        st.rerun()

st.subheader("📌 1. 請選擇案件類型與生成格式")
app_mode = st.radio(
    "請選擇要處理的意外類型：",
    [
        "TAPI (有人受傷 - 生成 Excel 白皮書)", 
        "TADO (僅車損 - 生成 Word 附錄)", 
        "T24 (正式報告 - 生成 T24 Word 檔)"
    ],
    horizontal=True
)

st.divider()

st.subheader("📸 2. 分類上載相片")
tab_main, tab_veh_a, tab_veh_b, tab_veh_c = st.tabs([
    "📋 手填紀錄單 (IMSP-T3) / 現場照片",
    "🚗 車輛 甲 相片",
    "🚙 車輛 乙 相片",
    "🚚 其他 / 車輛 丙 相片"
])

with tab_main:
    files_main = st.file_uploader("請上載【手填紀錄單 (IMSP-T3) / 現場照片 / 總表】", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="main_files")

with tab_veh_a:
    files_veh_a = st.file_uploader("請上載【車輛 甲】行車證及駕照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_a_files")

with tab_veh_b:
    files_veh_b = st.file_uploader("請上載【車輛 乙】行車證及駕照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_b_files")

with tab_veh_c:
    files_veh_c = st.file_uploader("請上載【車輛 丙 或其他】相片", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_c_files")

st.divider()

st.subheader("📝 3. 司機口供及補充文字 (可選填)")
statement = st.text_area("請輸入額外口供、酒精測試結果或補充資料", height=100)

def generate_tado_word(data, statement_text):
    if not docx: return None
    try:
        doc = docx.Document()
        h = doc.add_heading('交通意外 (TADO) 附錄報告', level=1)
        if hasattr(WD_ALIGN_PARAGRAPH, 'CENTER'): h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        vehicles = data.get('vehicles', [])
        veh_labels = ["甲", "乙", "丙", "丁", "戊", "己"]
        
        for idx, v in enumerate(vehicles):
            label_char = veh_labels[idx] if idx < len(veh_labels) else str(idx + 1)
            doc.add_heading(f"車輛{label_char}：", level=2)
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.2
            
            p.add_run("1）車牌： ").bold = True
            p.add_run(f"{v.get('plate_no', '———')}\n")
            
            v_class_en = v.get('vehicle_class_en', v.get('vehicle_class', '———'))
            v_class_zh = v.get('vehicle_class_zh', '———')
            v_type_str = f"{v_class_en} / {v_class_zh}" if v_class_zh != '———' else v_class_en
            p.add_run("2）車輛類型： ").bold = True
            p.add_run(f"{v_type_str}\n")
            
            p.add_run("3）總重量： ").bold = True
            p.add_run(f"{v.get('gross_weight', '———')}\n")
            p.add_run("4）廠名： ").bold = True
            p.add_run(f"{v.get('make', '———')}\n")
            p.add_run("5）顏色： ").bold = True
            p.add_run(f"{v.get('color', '———')}\n")
            p.add_run("6）存檔編號： ").bold = True
            p.add_run(f"{v.get('transaction_no', '———')}\n")
            p.add_run("7）屆滿日期： ").bold = True
            p.add_run(f"{v.get('expiry_date', '———')}\n")
            
            d_en = v.get('driver_name_en', '')
            d_zh = v.get('driver_name_zh', '')
            driver_str = f"{d_en} ({d_zh})" if d_zh else (d_en if d_en else '———')
            p.add_run("8）駕駛人： ").bold = True
            p.add_run(f"{driver_str}\n")
            
            p.add_run("9）駕駛執照號碼： ").bold = True
            p.add_run(f"{v.get('driving_license_no', '———')}\n")
            p.add_run("10）地址： ").bold = True
            p.add_run(f"\n   中文：{v.get('address_zh', '———')}\n   英文：{v.get('address_en', '———')}\n")
            p.add_run("11）電話號碼： ").bold = True
            p.add_run(f"{v.get('phone', '———')}\n")
            p.add_run("12）車輛損毀情況： ").bold = True
            p.add_run(f"{v.get('damage', '———')}\n")
            
            d_statement = v.get('driver_statement', statement_text if statement_text else '———')
            if d_statement and d_statement != '———':
                s_clean = d_statement.strip()
                d_statement_formatted = s_clean if (s_clean.startswith('「') and s_clean.endswith('」')) else f"「{s_clean}」"
            else:
                d_statement_formatted = '———'
                
            p.add_run("13）司機口供： ").bold = True
            p.add_run(f"{d_statement_formatted}\n")

        doc.add_heading("警方資料及測試紀錄：", level=2)
        p_police = doc.add_paragraph()
        p_police.paragraph_format.line_spacing = 1.2
        p_police.add_run("• 警車車牌： ").bold = True
        p_police.add_run(f"{data.get('police_veh_no', '———')}\n")
        p_police.add_run("• 到場警員： ").bold = True
        p_police.add_run(f"{data.get('police_no', '———')}\n")
        p_police.add_run("• 報案號碼 (RN)： ").bold = True
        p_police.add_run(f"{data.get('rn_no', '———')}\n")
        p_police.add_run("• 酒精呼氣測試 (PBT)： ").bold = True
        p_police.add_run(f"{data.get('breathalyzer_test', '未有提供 / 無紀錄')}\n")

        return doc
    except Exception as err:
        st.error(f"建立 Word 文件時發生錯誤：{str(err)}")
        return None

if "TAPI" in app_mode:
    button_label = "🚀 一鍵生成【Excel 白皮書 + Word 附錄】"
elif "TADO" in app_mode:
    button_label = "📄 一鍵生成【Word 附錄報告】"
else:
    button_label = "📝 一鍵生成【T24 正式 Word 報告 + Word 附錄】"

if st.button(button_label, type="primary"):
    images_with_tags = []
    
    def process_files(file_list, category_label):
        imgs = []
        if file_list:
            for f in file_list:
                img = Image.open(f)
                imgs.append((img, category_label))
        return imgs

    images_with_tags.extend(process_files(files_main, "手填紀錄單 (IMSP-T3)"))
    images_with_tags.extend(process_files(files_veh_a, "車輛 甲"))
    images_with_tags.extend(process_files(files_veh_b, "車輛 乙"))
    images_with_tags.extend(process_files(files_veh_c, "車輛 丙"))

    if not api_key:
        st.error("請先輸入或設定 Gemini API Key！")
    elif not images_with_tags and not statement:
        st.error("請最少上載一張相片或輸入文字資料！")
    else:
        with st.spinner('AI 正在全方位辨識並生成報告...'):
            try:
                client = genai.Client(api_key=api_key.strip())
                raw_images = [item[0] for item in images_with_tags]
                photo_notes = "\n".join([f"- 相片 {idx+1} 來自：【{item[1]}】" for idx, item in enumerate(images_with_tags)])
                
                prompt = f"""
                你是一個專業的香港交通意外資料提取助手。
                請仔細分析所有上載相片與現場文字內容：
                
                【相片來源分類】：
                {photo_notes}
                
                【補充文字內容與口供】：
                {statement}
                
                【📋 手填紀錄單 IMSP-T3 精準提取】：
                1. 意外時間 (ta_time，如 1542)
                2. vehicles 陣列：plate_no, vehicle_class_en, vehicle_class_zh
                3. injured_total 受傷人數與 injured_list 傷者資料
                4. tunnel_staff_arrived, police_arrived_time, police_no, police_veh_no, ai_staff_no, ai_arrived_time, ai_departed_time, rn_no, tunnel_staff_departed
                
                【🚘 車牌號碼】：完整提取含尾數！
                【🌐 地址】：同時提供繁體中文 address_zh 與英文 address_en！
                【🆔 證件】：hkid 尾數帶括號如 G111759(8)；driving_license_no 無括號！
                
                請嚴格以 JSON 格式輸出：
                {{
                    "log_no": "紀錄號碼",
                    "rn_no": "RN報案號碼",
                    "ta_date": "意外日期",
                    "ta_time": "意外時間",
                    "location": "意外地點",
                    "tunnel_staff_arrived": "到達時間",
                    "tunnel_staff_departed": "離開時間",
                    "police_arrived_time": "警察到達時間",
                    "police_no": "警員編號",
                    "ai_staff_no": "AI警員編號",
                    "ai_arrived_time": "AI到達時間",
                    "ai_departed_time": "AI離開時間",
                    "police_veh_no": "警車車牌",
                    "breathalyzer_test": "PBT紀錄",
                    "vehicles": [
                        {{
                            "plate_no": "車牌號碼",
                            "vehicle_class_en": "英文車類",
                            "vehicle_class_zh": "中文車類",
                            "gross_weight": "總重",
                            "make": "廠名",
                            "color": "顏色",
                            "transaction_no": "存檔號",
                            "expiry_date": "屆滿日",
                            "seating_capacity": "載客額",
                            "damage": "損毀情況",
                            "driver_name_en": "司機英文名",
                            "driver_name_zh": "司機中文名",
                            "age": "年齡",
                            "phone": "電話",
                            "hkid": "身份證",
                            "driving_license_no": "駕照號碼",
                            "dl_class": "駕照類別",
                            "dl_expiry_date": "駕照屆滿日",
                            "address_zh": "中文地址",
                            "address_en": "英文地址",
                            "driver_statement": "司機口供"
                        }}
                    ],
                    "injured_total": "受傷人數",
                    "injured_list": [
                        {{
                            "name_en": "英文名",
                            "name_zh": "中文名",
                            "age": "年齡",
                            "address_zh": "中文地址",
                            "address_en": "英文地址",
                            "phone": "電話",
                            "vehicle_no": "車牌",
                            "hospital": "醫院",
                            "escort": "備註"
                        }}
                    ]
                }}
                """
                
                contents = raw_images + [prompt]
                candidate_models = list(dict.fromkeys([selected_model, "gemini-1.5-flash", "gemini-2.5-pro"]))
                
                response = None
                last_exception = None

                for m_name in candidate_models:
                    try:
                        res = client.models.generate_content(
                            model=m_name,
                            contents=contents,
                            config=types.GenerateContentConfig(response_mime_type="application/json")
                        )
                        if res and res.text:
                            response = res
                            break
                    except Exception as e:
                        last_exception = e
                        time.sleep(1)
                        continue

                if not response or not response.text:
                    raise last_exception if last_exception else Exception("無法從 Google API 取得回應，請稍後再試。")
                
                data = json.loads(response.text)
                vehicles_list = data.get('vehicles', [])
                first_plate = vehicles_list[0].get('plate_no', 'Unknown').replace(' ', '_') if len(vehicles_list) > 0 else 'Unknown'
                
                date_std, date_zh, weekday_zh = parse_date_and_weekday(data.get('ta_date', ''))

                # 1. 生成 TADO Word 附錄 Bytes
                tado_bytes = None
                tado_filename = f"TADO_附錄報告_{first_plate}.docx"
                doc_tado = generate_tado_word(data, statement)
                if doc_tado:
                    tado_io = io.BytesIO()
                    doc_tado.save(tado_io)
                    tado_bytes = tado_io.getvalue()

                # 2. 生成 主報告 Bytes
                main_bytes = None
                main_filename = ""
                main_mime = ""
                t24_2_bytes = None
                t24_2_filename = ""

                if "TAPI" in app_mode:
                    wb = openpyxl.load_workbook("T26_白皮書_template_v2.xlsx")
                    ws_data = wb['Data']
                    ws_police = wb['Police']
                    
                    if data.get('log_no'): ws_data['B1'] = data.get('log_no')
                    if data.get('rn_no'): ws_data['B2'] = data.get('rn_no')
                    if data.get('ta_date'): ws_data['B3'] = date_std
                    if data.get('ta_time'): ws_data['B4'] = data.get('ta_time')
                    if data.get('location'): ws_data['B5'] = data.get('location')
                    if data.get('tunnel_staff_arrived'): ws_data['B6'] = data.get('tunnel_staff_arrived')
                    if data.get('tunnel_staff_departed'): ws_data['B7'] = data.get('tunnel_staff_departed')
                    if data.get('police_arrived_time'): ws_data['E2'] = data.get('police_arrived_time')
                    if data.get('police_no'): ws_data['E3'] = data.get('police_no')
                    if data.get('ai_staff_no'): ws_data['E4'] = data.get('ai_staff_no')
                    if data.get('ai_arrived_time'): ws_data['E5'] = data.get('ai_arrived_time')
                    if data.get('ai_departed_time'): ws_data['E6'] = data.get('ai_departed_time')
                    
                    v_mappings = [
                        {"plate": "B12", "veh_class": "B13", "make": "B14", "color": "B15", "trans_no": "B16", "expiry": "B17", "seating": "B18", "damage": "B20", "name_en": "E12", "name_zh": "E13", "age": "E14", "addr_zh": "E15", "addr_en": "E16", "phone": "E17", "hkid": "E18", "lic_no": "E19", "dl_class": "E20", "dl_exp": "E21"},
                        {"plate": "B27", "veh_class": "B28", "make": "B29", "color": "B30", "trans_no": "B31", "expiry": "B32", "seating": "B33", "damage": "B35", "name_en": "E27", "name_zh": "E28", "age": "E29", "addr_zh": "E30", "addr_en": "E31", "phone": "E32", "hkid": "E33", "lic_no": "E34", "dl_class": "E35", "dl_exp": "E36"},
                        {"plate": "B43", "veh_class": "B44", "make": "B45", "color": "B46", "trans_no": "B47", "expiry": "B48", "seating": "B49", "damage": "B51", "name_en": "E43", "name_zh": "E44", "age": "E45", "addr_zh": "E46", "addr_en": "E47", "phone": "E48", "hkid": "E49", "lic_no": "E50", "dl_class": "E51", "dl_exp": "E52"},
                        {"plate": "B58", "veh_class": "B59", "make": "B60", "color": "B61", "trans_no": "B62", "expiry": "B63", "seating": "B64", "damage": "B66", "name_en": "E58", "name_zh": "E59", "age": "E60", "addr_zh": "E61", "addr_en": "E62", "phone": "E63", "hkid": "E64", "lic_no": "E65", "dl_class": "E66", "dl_exp": "E67"}
                    ]
                    for idx, v in enumerate(vehicles_list):
                        if idx >= len(v_mappings): break
                        m = v_mappings[idx]
                        ws_data[m["plate"]] = v.get('plate_no', '')
                        if v.get('vehicle_class_en'): ws_data[m["veh_class"]] = v.get('vehicle_class_en')
                        ws_data[m["make"]] = v.get('make', '')
                        ws_data[m["color"]] = v.get('color', '')
                        if v.get('transaction_no'): ws_data[m["trans_no"]] = v.get('transaction_no')
                        if v.get('expiry_date'): ws_data[m["expiry"]] = v.get('expiry_date')
                        if v.get('seating_capacity'): ws_data[m["seating"]] = v.get('seating_capacity')
                        ws_data[m["damage"]] = v.get('damage', '')
                        ws_data[m["name_en"]] = v.get('driver_name_en', '')
                        ws_data[m["name_zh"]] = v.get('driver_name_zh', '')
                        if v.get('age'): ws_data[m["age"]] = v.get('age')
                        ws_data[m["addr_zh"]] = v.get('address_zh', '')
                        ws_data[m["addr_en"]] = v.get('address_en', '')
                        ws_data[m["phone"]] = v.get('phone', '')
                        hkid_val = format_hkid_parentheses(v.get('hkid', ''))
                        lic_val = v.get('driving_license_no', '')
                        if hkid_val: ws_data[m["hkid"]] = hkid_val
                        if lic_val: ws_data[m["lic_no"]] = lic_val
                        if v.get('dl_class'): ws_data[m["dl_class"]] = v.get('dl_class')
                        if v.get('dl_expiry_date'): ws_data[m["dl_exp"]] = v.get('dl_expiry_date')

                    if data.get('injured_total'): ws_data['B74'] = data.get('injured_total')
                    injured_list = data.get('injured_list', [])
                    i_mappings = [
                        {"en": "B77", "zh": "B78", "age": "B79", "azh": "B80", "aen": "B81", "ph": "B82", "vno": "B83", "hosp": "B84", "esc": "B85"},
                        {"en": "E77", "zh": "E78", "age": "E79", "azh": "E80", "aen": "E81", "ph": "E82", "vno": "E83", "hosp": "E84", "esc": "E85"},
                        {"en": "B89", "zh": "B90", "age": "B91", "azh": "B92", "aen": "B93", "ph": "B94", "vno": "B95", "hosp": "B96", "esc": "B97"},
                        {"en": "E89", "zh": "E90", "age": "E91", "azh": "E92", "aen": "E93", "ph": "E94", "vno": "E95", "hosp": "E96", "esc": "E97"}
                    ]
                    for idx, inj in enumerate(injured_list):
                        if idx >= len(i_mappings): break
                        im = i_mappings[idx]
                        ws_data[im["en"]] = inj.get('name_en', '')
                        ws_data[im["zh"]] = inj.get('name_zh', '')
                        if inj.get('age'): ws_data[im["age"]] = inj.get('age')
                        ws_data[im["azh"]] = inj.get('address_zh', '')
                        ws_data[im["aen"]] = inj.get('address_en', '')
                        ws_data[im["ph"]] = inj.get('phone', '')
                        ws_data[im["vno"]] = inj.get('vehicle_no', '')
                        ws_data[im["hosp"]] = inj.get('hospital', '')
                        ws_data[im["esc"]] = inj.get('escort', '')

                    if statement: ws_police.cell(row=5, column=1, value=f"現場紀錄及司機口供：\n{statement}")
                    if os.path.exists("car_diagram.png"):
                        if 'P.1' in wb.sheetnames:
                            ws_p1 = wb['P.1']
                            img_p1 = OpenpyxlImage("car_diagram.png")
                            img_p1.width, img_p1.height = 414, 347
                            ws_p1.add_image(img_p1, 'I41')
                        if 'P.2' in wb.sheetnames:
                            ws_p2 = wb['P.2']
                            img_p2_1 = OpenpyxlImage("car_diagram.png")
                            img_p2_1.width, img_p2_1.height = 434, 355
                            ws_p2.add_image(img_p2_1, 'H14')
                            img_p2_2 = OpenpyxlImage("car_diagram.png")
                            img_p2_2.width, img_p2_2.height = 437, 334
                            ws_p2.add_image(img_p2_2, 'H39')
                    
                    out_io = io.BytesIO()
                    wb.save(out_io)
                    main_bytes = out_io.getvalue()
                    main_filename = f"白皮書_{first_plate}.xlsx"
                    main_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                elif "T24" in app_mode:
                    context = {
                        "{{RN_NO}}": data.get('rn_no', '———'),
                        "{{TA_DATE}}": date_std,
                        "{{TA_DATE_ZH}}": date_zh,
                        "{{TA_TIME}}": data.get('ta_time', '———'),
                        "{{TA_WEEKDAY}}": weekday_zh,
                        "{{LOCATION}}": data.get('location', '———'),
                        "{{POLICE_NO}}": data.get('police_no', '———'),
                        "{{POLICE_VEH_NO}}": data.get('police_veh_no', '———'),
                        "{{PBT}}": data.get('breathalyzer_test', '———')
                    }
                    for i in range(1, 7):
                        v_idx = i - 1
                        if v_idx < len(vehicles_list):
                            v = vehicles_list[v_idx]
                            d_statement = v.get('driver_statement', statement if statement else '———')
                            s_clean = d_statement.strip() if d_statement and d_statement != '———' else '———'
                            d_statement_formatted = s_clean if (s_clean.startswith('「') and s_clean.endswith('」')) else (f"「{s_clean}」" if s_clean != '———' else '———')

                            context["{{" + f"VEH_{i}_PLATE" + "}}"] = v.get('plate_no', '———')
                            context["{{" + f"VEH_{i}_CLASS_ZH" + "}}"] = v.get('vehicle_class_zh', v.get('vehicle_class_en', '———'))
                            context["{{" + f"VEH_{i}_CLASS_EN" + "}}"] = v.get('vehicle_class_en', '———')
                            context["{{" + f"VEH_{i}_WEIGHT" + "}}"] = v.get('gross_weight', '———')
                            context["{{" + f"VEH_{i}_MAKE" + "}}"] = v.get('make', '———')
                            context["{{" + f"VEH_{i}_COLOR" + "}}"] = v.get('color', '———')
                            context["{{" + f"VEH_{i}_TRANS_NO" + "}}"] = v.get('transaction_no', '———')
                            context["{{" + f"VEH_{i}_EXPIRY" + "}}"] = v.get('expiry_date', '———')
                            context["{{" + f"VEH_{i}_DAMAGE" + "}}"] = v.get('damage', '———')
                            context["{{" + f"DRV_{i}_NAME_ZH" + "}}"] = v.get('driver_name_zh', '———')
                            context["{{" + f"DRV_{i}_NAME_EN" + "}}"] = v.get('driver_name_en', '———')
                            context["{{" + f"DRV_{i}_HKID" + "}}"] = format_hkid_parentheses(v.get('hkid', '———'))
                            context["{{" + f"DRV_{i}_LIC" + "}}"] = v.get('driving_license_no', '———')
                            context["{{" + f"DRV_{i}_ADDR_ZH" + "}}"] = v.get('address_zh', '———')
                            context["{{" + f"DRV_{i}_ADDR_EN" + "}}"] = v.get('address_en', '———')
                            context["{{" + f"DRV_{i}_PHONE" + "}}"] = v.get('phone', '———')
                            context["{{" + f"DRV_{i}_STATEMENT" + "}}"] = d_statement_formatted
                        else:
                            for key_suffix in ["PLATE", "CLASS_ZH", "CLASS_EN", "WEIGHT", "MAKE", "COLOR", "TRANS_NO", "EXPIRY", "DAMAGE"]:
                                context["{{" + f"VEH_{i}_{key_suffix}" + "}}"] = "———"
                            for key_suffix in ["NAME_ZH", "NAME_EN", "HKID", "LIC", "ADDR_ZH", "ADDR_EN", "PHONE", "STATEMENT"]:
                                context["{{" + f"DRV_{i}_{key_suffix}" + "}}"] = "———"

                    doc1 = docx.Document("T24_template_1.docx")
                    replace_docx_tags(doc1, context)
                    out_io1 = io.BytesIO()
                    doc1.save(out_io1)
                    main_bytes = out_io1.getvalue()
                    main_filename = f"T24_正式報告_{first_plate}.docx"
                    main_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                    if len(vehicles_list) > 3 and os.path.exists("T24_template_2.docx"):
                        doc2 = docx.Document("T24_template_2.docx")
                        replace_docx_tags(doc2, context)
                        out_io2 = io.BytesIO()
                        doc2.save(out_io2)
                        t24_2_bytes = out_io2.getvalue()
                        t24_2_filename = f"T24_正式報告_(續頁)_{first_plate}.docx"

                # 3. 建立一鍵 ZIP 打包檔 (專為手機用戶設計)
                zip_io = io.BytesIO()
                with zipfile.ZipFile(zip_io, 'w', zipfile.ZIP_DEFLATED) as zf:
                    if main_bytes:
                        zf.writestr(main_filename, main_bytes)
                    if tado_bytes and "TADO" not in app_mode:
                        zf.writestr(tado_filename, tado_bytes)
                    if t24_2_bytes:
                        zf.writestr(t24_2_filename, t24_2_bytes)
                
                zip_bytes = zip_io.getvalue()
                zip_filename = f"交通意外報告打包_{first_plate}.zip"

                wa_text = f"🚓【交通意外報告已生成】\n• RN編號：{data.get('rn_no', '無')}\n• 意外時間：{data.get('ta_time', '無')}\n• 涉及車牌：{first_plate}\n• 模式：{app_mode.split(' ')[0]}\n請在附件接收報告檔。"
                
                st.session_state["current_mode"] = app_mode
                st.session_state["main_bytes"] = main_bytes
                st.session_state["main_filename"] = main_filename
                st.session_state["main_mime"] = main_mime
                st.session_state["tado_bytes"] = tado_bytes
                st.session_state["tado_filename"] = tado_filename
                st.session_state["t24_2_bytes"] = t24_2_bytes
                st.session_state["t24_2_filename"] = t24_2_filename
                st.session_state["zip_bytes"] = zip_bytes
                st.session_state["zip_filename"] = zip_filename
                st.session_state["wa_url"] = f"https://api.whatsapp.com/send?text={urllib.parse.quote(wa_text)}"
                st.session_state["preview_data"] = data
                st.session_state["generated"] = True
                
                st.rerun()

            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")

# ----------------- 4. 下載區 -----------------
if st.session_state.get("generated"):
    st.subheader("📊 AI 自動提取與精準對齊預覽：")
    st.json(st.session_state["preview_data"])
    
    st.success("🎉 報告生成成功！手機用戶建議點擊【📦 一鍵打包下載】，一次過將所有報告存入手機：")
    
    cur_mode = st.session_state.get("current_mode", "")
    
    # 🌟 iPhone 手機救星：一鍵 ZIP 打包下載按鈕
    if st.session_state.get("zip_bytes") and ("TAPI" in cur_mode or "T24" in cur_mode):
        st.download_button(
            label=f"📦 1. 一鍵打包下載所有報告 (.zip 檔 - 專為手機免跳頁設計)",
            data=st.session_state["zip_bytes"],
            file_name=st.session_state["zip_filename"],
            mime="application/zip",
            key="dl_zip_btn"
        )
        st.write("---")

    # 個別下載按鈕區
    if "TAPI" in cur_mode or "T24" in cur_mode:
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label=f"📥 2a. 下載主要報告 ({st.session_state['main_filename']})",
                data=st.session_state["main_bytes"],
                file_name=st.session_state["main_filename"],
                mime=st.session_state["main_mime"],
                key="dl_main_btn"
            )
        with col_dl2:
            if st.session_state["tado_bytes"]:
                st.download_button(
                    label=f"📄 2b. 下載 Word 附錄報告 ({st.session_state['tado_filename']})",
                    data=st.session_state["tado_bytes"],
                    file_name=st.session_state["tado_filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_tado_btn"
                )

        if st.session_state.get("t24_2_bytes"):
            st.download_button(
                label=f"📥 下載 T24 正式報告續頁 ({st.session_state['t24_2_filename']})",
                data=st.session_state["t24_2_bytes"],
                file_name=st.session_state["t24_2_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_t24_2_btn"
            )

    else:
        if st.session_state["tado_bytes"]:
            st.download_button(
                label=f"📄 下載 Word 附錄報告 ({st.session_state['tado_filename']})",
                data=st.session_state["tado_bytes"],
                file_name=st.session_state["tado_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_tado_single_btn"
            )

    st.markdown(f'<a href="{st.session_state["wa_url"]}" target="_blank" class="wa-button">💬 分享案件通報至 WhatsApp 群組</a>', unsafe_allow_html=True)
    
    st.divider()
    if st.button("🔄 重新開始 / 處理下宗案件"):
        for k in state_keys:
            st.session_state[k] = state_keys[k]
        st.rerun()
