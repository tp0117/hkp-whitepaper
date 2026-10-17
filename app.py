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

# 1. 網頁標題及頁面設定 (必須是第一個 Streamlit 指令)
st.set_page_config(page_title="交通意外資料處理系統 (TAPI / TADO / T24)", page_icon="🚓", layout="wide")

# 2. UI 手機化美化 (Custom CSS)
st.markdown("""
    <style>
    /* 隱藏 Streamlit 原生頂部選單、Header 及 Footer */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* 調整主容器邊距，更貼合手機螢幕 */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    
    /* 放大生成與重置按鈕，適合手機大拇指點擊 */
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
    
    /* WhatsApp 分享按鈕專用綠色樣式 */
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
    
    /* 修復深色模式 (Dark Mode) 下單選框文字變白看不清的問題 */
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

# Session State 狀態初始化 (解決點擊下載後畫面刷新導致按鈕消失的問題)
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
    "wa_url": "",
    "preview_data": None
}
for key, default_val in state_keys.items():
    if key not in st.session_state:
        st.session_state[key] = default_val

# Helper 函數：僅用於身份證號碼尾數加上圓括號，如 G111759(8)
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

# Helper 函數：替換 Word 文件中的所有 Tag
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

# 🔒 密碼鎖設定：優先從 Secrets 讀取 APP_PASSWORD，若未設定則預設為 "8888"
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

# ----------------- 以下為驗證通過後的系統主介面 -----------------

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
    api_key = st.text_input("Gemini API Key", value=default_api_key, type="password", help="系統已自動讀取雲端預設 Key，亦可手動覆蓋")
    st.markdown("[按此免費獲取 Google Gemini API Key](https://aistudio.google.com/)")
    
    st.divider()
    selected_model = st.selectbox(
        "🤖 首選 AI 模型",
        ["gemini-3.8-flash", "gemini-1.5-flash", "gemini-2.5-pro"],
        index=0,
        help="若首選模型額度滿，系統將自動無縫嘗試其他可用模型"
    )
    
    st.divider()
    if st.button("🔒 登出系統"):
        st.session_state["authenticated"] = False
        st.rerun()

# ----------------- 1. 選擇案件處理模式 -----------------
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

# ----------------- 2. 分區上載相片 -----------------
st.subheader("📸 2. 分類上載相片 (支援行車證、駕照、手填紀錄單 IMSP-T3、現場照片)")

tab_main, tab_veh_a, tab_veh_b, tab_veh_c = st.tabs([
    "📋 手填紀錄單 (IMSP-T3) / 現場照片",
    "🚗 車輛 甲 (Vehicle A) 相片",
    "🚙 車輛 乙 (Vehicle B) 相片",
    "🚚 其他 / 車輛 丙 (Vehicle C) 相片"
])

with tab_main:
    files_main = st.file_uploader("請上載【手填紀錄單 (IMSP-T3) / 現場照片 / 總表】", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="main_files")

with tab_veh_a:
    files_veh_a = st.file_uploader("請上載【車輛 甲 (Vehicle A)】的行車證及司機駕駛執照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_a_files")

with tab_veh_b:
    files_veh_b = st.file_uploader("請上載【車輛 乙 (Vehicle B)】的行車證及司機駕駛執照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_b_files")

with tab_veh_c:
    files_veh_c = st.file_uploader("請上載【車輛 丙 (Vehicle C) 或其他】相片", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_c_files")

st.divider()

# ----------------- 3. 補充口供 -----------------
st.subheader("📝 3. 司機口供及補充文字 (可選填)")
statement = st.text_area("請輸入額外口供、酒精測試結果或補充資料 (若相片已包含所有資料，此處可留空)", height=100)

# Helper: 生成 TADO Word 附錄檔
def generate_tado_word(data, statement_text):
    if not docx:
        return None
    try:
        doc = docx.Document()
        h = doc.add_heading('交通意外 (TADO) 附錄報告', level=1)
        if hasattr(WD_ALIGN_PARAGRAPH, 'CENTER'):
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        vehicles = data.get('vehicles', [])
        veh_labels = ["甲", "乙", "丙", "丁", "戊", "己"]
        
        for idx, v in enumerate(vehicles):
            label_char = veh_labels[idx] if idx < len(veh_labels) else str(idx + 1)
            doc.add_heading(f"車輛{label_char}：", level=2)
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.2
            
            plate = v.get('plate_no', '———')
            p.add_run("1）車牌： ").bold = True
            p.add_run(f"{plate}\n")
            
            v_class_en = v.get('vehicle_class_en', v.get('vehicle_class', '———'))
            v_class_zh = v.get('vehicle_class_zh', '———')
            v_type_str = f"{v_class_en} / {v_class_zh}" if v_class_zh != '———' else v_class_en
            p.add_run("2）車輛類型： ").bold = True
            p.add_run(f"{v_type_str}\n")
            
            weight = v.get('gross_weight', '———')
            p.add_run("3）總重量： ").bold = True
            p.add_run(f"{weight}\n")
            
            make = v.get('make', '———')
            p.add_run("4）廠名： ").bold = True
            p.add_run(f"{make}\n")
            
            color = v.get('color', '———')
            p.add_run("5）顏色： ").bold = True
            p.add_run(f"{color}\n")
            
            trans_no = v.get('transaction_no', '———')
            p.add_run("6）存檔編號： ").bold = True
            p.add_run(f"{trans_no}\n")
            
            exp = v.get('expiry_date', '———')
            p.add_run("7）屆滿日期： ").bold = True
            p.add_run(f"{exp}\n")
            
            d_en = v.get('driver_name_en', '')
            d_zh = v.get('driver_name_zh', '')
            driver_str = f"{d_en} ({d_zh})" if d_zh else (d_en if d_en else '———')
            p.add_run("8）駕駛人： ").bold = True
            p.add_run(f"{driver_str}\n")
            
            lic_no = v.get('driving_license_no', '———')
            p.add_run("9）駕駛執照號碼： ").bold = True
            p.add_run(f"{lic_no}\n")
            
            addr_zh = v.get('address_zh', '———')
            addr_en = v.get('address_en', '———')
            p.add_run("10）地址： ").bold = True
            p.add_run(f"\n   中文：{addr_zh}\n   英文：{addr_en}\n")
            
            phone = v.get('phone', '———')
            p.add_run("11）電話號碼： ").bold = True
            p.add_run(f"{phone}\n")
            
            damage = v.get('damage', '———')
            p.add_run("12）車輛損毀情況： ").bold = True
            p.add_run(f"{damage}\n")
            
            d_statement = v.get('driver_statement', statement_text if statement_text else '———')
            if d_statement and d_statement != '———':
                s_clean = d_statement.strip()
                if not (s_clean.startswith('「') and s_clean.endswith('」')):
                    d_statement_formatted = f"「{s_clean}」"
                else:
                    d_statement_formatted = s_clean
            else:
                d_statement_formatted = '———'
                
            p.add_run("13）司機口供： ").bold = True
            p.add_run(f"{d_statement_formatted}\n")

        doc.add_heading("警方資料及測試紀錄：", level=2)
        p_police = doc.add_paragraph()
        p_police.paragraph_format.line_spacing = 1.2
        
        police_veh = data.get('police_veh_no', '———')
        police_officer = data.get('police_no', '———')
        rn_no = data.get('rn_no', '———')
        pbt = data.get('breathalyzer_test', '未有提供 / 無紀錄')
        
        p_police.add_run("• 警車車牌： ").bold = True
        p_police.add_run(f"{police_veh}\n")
        p_police.add_run("• 到場警員： ").bold = True
        p_police.add_run(f"{police_officer}\n")
        p_police.add_run("• 報案號碼 (RN)： ").bold = True
        p_police.add_run(f"{rn_no}\n")
        p_police.add_run("• 酒精呼氣測試 (PBT)： ").bold = True
        p_police.add_run(f"{pbt}\n")

        return doc
    except Exception as err:
        st.error(f"建立 Word 文件時發生錯誤：{str(err)}")
        return None

# 動態生成按鈕標題
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
    images_with_tags.extend(process_files(files_veh_a, "車輛 甲 (Vehicle A)"))
    images_with_tags.extend(process_files(files_veh_b, "車輛 乙 (Vehicle B)"))
    images_with_tags.extend(process_files(files_veh_c, "車輛 丙 (Vehicle C)"))

    if not api_key:
        st.error("請先輸入或設定 Gemini API Key！")
    elif not images_with_tags and not statement:
        st.error("請最少上載一張相片或輸入文字資料！")
    else:
        with st.spinner('AI 正在全方位辨識（包含對齊 IMSP-T3 頂部車輛類別與傷者資訊）...'):
            try:
                client = genai.Client(api_key=api_key.strip())
                raw_images = [item[0] for item in images_with_tags]
                photo_notes = "\n".join([f"- 相片 {idx+1} 來自：【{item[1]}】" for idx, item in enumerate(images_with_tags)])
                
                prompt = f"""
                你是一個專業的香港交通意外資料提取助手。
                請仔細分析所有上載的相片（特別是東隧事故處理程序 IMSP-T3 手填紙本紀錄、行車證、駕駛執照、現場照片）以及現場文字內容：
                
                【相片來源分類】：
                {photo_notes}
                
                【補充文字內容與口供】：
                {statement}
                
                【📋 手填紀錄單 IMSP-T3 頂部與表格欄位精準提取加強對應】：
                1. 頂部「事故發生時間」：提取意外時間 (ta_time，如 1542)。
                2. 頂部「車輛：」欄位 (如 "P/C 8161 / M/C WW 5031")：
                   - P/C 代表 PRIVATE CAR (私家車)
                   - M/C 代表 MOTORCYCLE (電單車)
                   - L/G/V 或 LGV 代表 LIGHT GOODS VEHICLE (輕型貨車)
                   - H/G/V 代表 HEAVY GOODS VEHICLE (重型貨車)
                   - 請將提取出的車牌與車輛類型對應補入 vehicles 陣列的 plate_no, vehicle_class_en, vehicle_class_zh！
                3. 頂部「受傷人數：」及上方手寫註記 (如 "M/C WW5031 司機" 或 "P/C 8161 乘客")：
                   - injured_total: 受傷人數 (如 1)
                   - 請解析手寫文字中的涉及車牌 (如 WW5031) 與受傷者身份，寫入 injured_list 傷者資料 (vehicle_no, escort/notes)！
                4. 表格項目 1「東隧職員發現/到達現場時間」: tunnel_staff_arrived
                5. 表格項目 8「警方到達時間」: police_arrived_time
                6. 表格項目 8「警員編號」: police_no (普通軍裝警員編號，如 19626)
                7. 表格項目 8 或 11「車牌」: police_veh_no (警車車牌，如 AM8049 / AM2795)
                8. 表格項目 11「警方 AI 到達時間」: ai_arrived_time
                9. 表格項目 11「警方 AI 警員編號」: ai_staff_no (如 PC14367)
                10. 表格項目 11「警方 AI 離開時間」: ai_departed_time
                11. 表格項目 12「警方檔案編號」: rn_no (如 RN=13913)
                12. 表格項目 19「所有肇事車輛及政府人員離開」: tunnel_staff_departed (如 1648)
                
                【🚘 車牌號碼強制要求】：
                - plate_no：必須完整提取並包含車牌的尾數字母/數字 (例如 UG8125 R, WA1062 Q, WW5031, AM8049)。絕對不可以省略尾數！
                
                【🌐 地址雙向翻譯強制要求 (100% 必須完成)】：
                1. 司機及傷者的住址：相片中不論提供的是中文住址還是英文住址，你必須同時生成：
                   - address_zh：完整繁體中文住址 (若原本是英文，必須精準翻譯成繁體中文)。
                   - address_en：完整英文住址 (若原本是中文，必須精準翻譯成英文)。
                2. 絕對不可留空或填寫 N/A！兩者必須同時填滿。
                
                【🆔 證件號碼與期滿日期要求】：
                1. 身份證號碼 (hkid)：最後一個校驗碼數字或字母必須使用圓括號 ( ) 括住，例如 G111759(8) 或 Y463269(7)。
                2. 駕駛執照號碼 (driving_license_no)：請照實提取，尾數【絕對不要添加括號】(例如 Y463269 7)。
                3. dl_expiry_date (駕駛執照期滿日期)：駕駛執照上的有效期/屆滿日期 (如 DD/MM/YYYY 或 DD-MM-YYYY)。
                
                【受傷者 (Injured) 詳細資料與電話號碼提取】：
                1. injured_total: 受傷者合共人數 (無人受傷填 "0")。
                2. injured_list: 傷者列表陣列 (最多 4 人)，必須完整提取：
                   - name_en: 傷者英文姓名
                   - name_zh: 傷者中文姓名
                   - age: 年齡
                   - address_zh: 完整中文住址 (需雙向翻譯)
                   - address_en: 完整英文住址 (需雙向翻譯)
                   - phone: 傷者電話號碼 (必須照實提取，不可遺漏)
                   - vehicle_no: 涉及車牌號碼 (從 IMSP-T3 頂部標記如 WW5031 提取)
                   - hospital: 送往醫院
                   - escort: 身份備註 (如 M/C WW5031 司機)
                
                請嚴格以 JSON 格式輸出：
                {{
                    "log_no": "Log No 紀錄號碼",
                    "rn_no": "HKP RN 報案號碼/案號 (如 RN:13913)",
                    "ta_date": "意外日期",
                    "ta_time": "意外時間 (如 1542)",
                    "location": "意外地點",
                    "tunnel_staff_arrived": "隧道職員到達時間 (如 1542)",
                    "tunnel_staff_departed": "隧道職員離開時間 (如 1648)",
                    "police_arrived_time": "警察到達時間 (如 1604)",
                    "police_no": "普通警員編號 (如 19626)",
                    "ai_staff_no": "交通意外調查組人員編號 (如 PC14367)",
                    "ai_arrived_time": "HKP AI 到達時間 (如 1616)",
                    "ai_departed_time": "HKP AI 離開時間 (如 1640)",
                    "police_veh_no": "警車車牌 (如 AM8049)",
                    "breathalyzer_test": "酒精呼氣測試紀錄",
                    "vehicles": [
                        {{
                            "plate_no": "車牌號碼 (必須含尾數，如 WW5031 或 8161)",
                            "vehicle_class_en": "英文車輛類別 (如 MOTORCYCLE)",
                            "vehicle_class_zh": "中文車輛類別 (如 電單車)",
                            "gross_weight": "總重量",
                            "make": "車輛廠名 (如 TOYOTA)",
                            "color": "車輛顏色",
                            "transaction_no": "存檔編號",
                            "expiry_date": "屆滿日期 (無則 ———)",
                            "seating_capacity": "載客額",
                            "damage": "車輛損毀情況",
                            "driver_name_en": "司機英文姓名",
                            "driver_name_zh": "司機中文姓名",
                            "age": "年齡",
                            "phone": "電話號碼",
                            "hkid": "身份證號碼 (尾數帶括號，如 G111759(8))",
                            "driving_license_no": "駕駛執照號碼 (無括號，如 Y463269 7)",
                            "dl_class": "駕照 Full Class (如 1, 2, 3)",
                            "dl_expiry_date": "駕駛執照期滿日期 (如 DD/MM/YYYY)",
                            "address_zh": "完整中文住址",
                            "address_en": "完整英文住址",
                            "driver_statement": "司機口供"
                        }}
                    ],
                    "injured_total": "受傷者合共人數 (如 1)",
                    "injured_list": [
                        {{
                            "name_en": "傷者英文名",
                            "name_zh": "傷者中文名",
                            "age": "年齡",
                            "address_zh": "中文住址",
                            "address_en": "英文住址",
                            "phone": "電話號碼",
                            "vehicle_no": "涉及車牌號碼 (如 WW5031)",
                            "hospital": "送往醫院",
                            "escort": "身份備註 (如 M/C WW5031 司機)"
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

                # ----------------- 1. 無論什麼模式，一律背景同步生成 TADO Word 附錄 -----------------
                tado_bytes = None
                tado_filename = f"TADO_附錄報告_{first_plate}.docx"
                doc_tado = generate_tado_word(data, statement)
                if doc_tado:
                    tado_io = io.BytesIO()
                    doc_tado.save(tado_io)
                    tado_bytes = tado_io.getvalue()

                # ----------------- 2. 根據選擇模式生成主報告 -----------------
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

                # ----------------- 3. 儲存至 Session State 以保證畫面按鈕持久化 -----------------
                wa_text = f"🚓【交通意外報告已生成】\n• RN編號：{data.get('rn_no', '無')}\n• 意外時間：{data.get('ta_time', '無')}\n• 涉及車牌：{first_plate}\n• 模式：{app_mode.split(' ')[0]}\n請在附件接收報告檔。"
                
                st.session_state["current_mode"] = app_mode
                st.session_state["main_bytes"] = main_bytes
                st.session_state["main_filename"] = main_filename
                st.session_state["main_mime"] = main_mime
                st.session_state["tado_bytes"] = tado_bytes
                st.session_state["tado_filename"] = tado_filename
                st.session_state["t24_2_bytes"] = t24_2_bytes
                st.session_state["t24_2_filename"] = t24_2_filename
                st.session_state["wa_url"] = f"https://api.whatsapp.com/send?text={urllib.parse.quote(wa_text)}"
                st.session_state["preview_data"] = data
                st.session_state["generated"] = True
                
                st.rerun()

            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")

# ----------------- 4. 持久化雙下載按鈕區 -----------------
if st.session_state.get("generated"):
    st.subheader("📊 AI 自動提取與精準對齊預覽：")
    st.json(st.session_state["preview_data"])
    
    st.success("🎉 報告生成成功！請於下方按鈕下載（點擊其中一個下載後，頁面按鈕會保持，可繼續點擊下載另一個）：")
    
    cur_mode = st.session_state.get("current_mode", "")
    
    # 若為 TAPI 或 T24 模式：顯示「主報告 + Word 附錄」雙按鈕
    if "TAPI" in cur_mode or "T24" in cur_mode:
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label=f"📥 1. 下載主要報告 ({st.session_state['main_filename']})",
                data=st.session_state["main_bytes"],
                file_name=st.session_state["main_filename"],
                mime=st.session_state["main_mime"],
                key="dl_main_btn"
            )
        with col_dl2:
            if st.session_state["tado_bytes"]:
                st.download_button(
                    label=f"📄 2. 下載 Word 附錄報告 ({st.session_state['tado_filename']})",
                    data=st.session_state["tado_bytes"],
                    file_name=st.session_state["tado_filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_tado_btn"
                )

        if st.session_state.get("t24_2_bytes"):
            st.download_button(
                label=f"📥 1b. 下載 T24 正式報告續頁 ({st.session_state['t24_2_filename']})",
                data=st.session_state["t24_2_bytes"],
                file_name=st.session_state["t24_2_filename"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_t24_2_btn"
            )

    # 若為 TADO 模式：單獨顯示 Word 附錄
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
