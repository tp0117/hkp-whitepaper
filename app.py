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
st.set_page_config(page_title="交通意外資料處理系統 (TAPI / TADO / T24 + T17A)", page_icon="🚓", layout="wide")

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

# Session State 狀態初始化 (解決點擊下載後頁面刷新遺失檔案問題)
state_keys = {
    "generated": False,
    "main_bytes": None,
    "main_filename": "",
    "main_mime": "",
    "t17_bytes": None,
    "t17_filename": "",
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

# Helper 函數：生成 IMSP-T17A Excel 相片報告表
def generate_t17_excel(template_path, data, photo_items, current_app_mode):
    if not os.path.exists(template_path):
        return None

    wb = openpyxl.load_workbook(template_path)
    total_photos = len(photo_items)
    sheets_needed = max(1, (total_photos + 1) // 2)

    date_std, _, _ = parse_date_and_weekday(data.get('ta_date', ''))
    date_val = date_std if date_std != '———' else data.get('ta_date', '')
    time_val = data.get('ta_time', '')
    
    topic_val = "T.A.P.I." if "TAPI" in current_app_mode else "T.A.D.O. (HIT&RUN)"
    location_val = data.get('location', '')
    officer_val = data.get('ai_staff_no') or data.get('police_no') or ''

    photo_idx = 0

    for sheet_num in range(1, 9):
        possible_names = [f"photo ({sheet_num})", f"PHOTO ({sheet_num})", f"photo({sheet_num})", f"PHOTO({sheet_num})"]
        s_name = None
        for name in possible_names:
            if name in wb.sheetnames:
                s_name = name
                break
        
        if not s_name:
            continue

        ws = wb[s_name]

        if sheet_num > sheets_needed:
            ws.sheet_state = 'hidden'
            continue

        if date_val and date_val != '———': ws['B44'] = date_val
        if time_val and time_val != '———': ws['F44'] = time_val
        ws['M44'] = topic_val
        if location_val and location_val != '———': ws['C46'] = location_val
        if officer_val and officer_val != '———': ws['N46'] = officer_val

        # 相片 (1) 定位 C5 / 圖說 C21
        if photo_idx < total_photos:
            p_item = photo_items[photo_idx]
            img_obj = p_item['img'].copy()
            img_obj.thumbnail((450, 260))
            
            img_io = io.BytesIO()
            img_obj.save(img_io, format='PNG')
            img_io.seek(0)
            
            xl_img = OpenpyxlImage(img_io)
            ws.add_image(xl_img, 'C5')

            if p_item.get('is_scene', False):
                ws['C21'] = "相片顯示肇事交通意外現場情況。"
            else:
                v_label = p_item.get('veh_label', '甲')
                v_class = p_item.get('veh_class_zh', '車輛')
                if not v_class or v_class == '———': v_class = '車輛'
                v_plate = p_item.get('plate_no', '')
                if v_plate == '———': v_plate = ''
                v_dmg = p_item.get('damage', '')
                if v_dmg == '———': v_dmg = '車身毀爛'
                ws['C21'] = f"相片顯示（車輛{v_label}）{v_class} {v_plate} {v_dmg}。"

            photo_idx += 1

        # 相片 (2) 定位 C24 / 圖說 C41
        if photo_idx < total_photos:
            p_item = photo_items[photo_idx]
            img_obj = p_item['img'].copy()
            img_obj.thumbnail((450, 260))
            
            img_io = io.BytesIO()
            img_obj.save(img_io, format='PNG')
            img_io.seek(0)
            
            xl_img = OpenpyxlImage(img_io)
            ws.add_image(xl_img, 'C24')

            if p_item.get('is_scene', False):
                ws['C41'] = "相片顯示肇事交通意外現場情況。"
            else:
                v_label = p_item.get('veh_label', '甲')
                v_class = p_item.get('veh_class_zh', '車輛')
                if not v_class or v_class == '———': v_class = '車輛'
                v_plate = p_item.get('plate_no', '')
                if v_plate == '———': v_plate = ''
                v_dmg = p_item.get('damage', '')
                if v_dmg == '———': v_dmg = '車身毀爛'
                ws['C41'] = f"相片顯示（車輛{v_label}）{v_class} {v_plate} {v_dmg}。"

            photo_idx += 1

    out_stream = io.BytesIO()
    wb.save(out_stream)
    return out_stream.getvalue()

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

st.title("🚓 交通意外資料處理系統 (TAPI / TADO / T24 + T17A)")

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

st.subheader("📌 1. 請選擇案件類型與生成格式 (預設連同 T17A 相片表雙檔輸出)")
app_mode = st.radio(
    "請選擇要處理的意外類型：",
    [
        "TAPI + T17 (有人受傷 - 生成 Excel 白皮書 + 相片表)", 
        "TADO + T17 (僅車損 - 生成 Word 附錄 + 相片表)", 
        "T24 + T17 (正式報告 - 生成 T24 Word 檔 + 相片表)"
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
    files_main = st.file_uploader("請上載【手填紀錄單 (IMSP-T3) / 現場照片】", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="main_files")

with tab_veh_a:
    files_veh_a = st.file_uploader("請上載【車輛 甲】行車證及駕照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_a_files")

with tab_veh_b:
    files_veh_b = st.file_uploader("請上載【車輛 乙】行車證及駕照", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_b_files")

with tab_veh_c:
    files_veh_c = st.file_uploader("請上載【車輛 丙 或其他】相片", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="veh_c_files")

st.divider()

st.subheader("📝 3. 司機口供及補充文字 (可選填)")
statement = st.text_area("請輸入額外口供或補充資料", height=100)

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

button_label = "🚀 立即生成雙份報告"

if st.button(button_label, type="primary"):
    images_with_tags = []
    photo_items_for_t17 = []
    
    def process_files(file_list, category_label, is_scene_flag, veh_label_char):
        imgs = []
        if file_list:
            for f in file_list:
                img = Image.open(f)
                imgs.append((img, category_label))
                photo_items_for_t17.append({
                    "img": img,
                    "is_scene": is_scene_flag,
                    "veh_label": veh_label_char
                })
        return imgs

    images_with_tags.extend(process_files(files_main, "手填紀錄單 (IMSP-T3)", True, "現場"))
    images_with_tags.extend(process_files(files_veh_a, "車輛 甲", False, "甲"))
    images_with_tags.extend(process_files(files_veh_b, "車輛 乙", False, "乙"))
    images_with_tags.extend(process_files(files_veh_c, "車輛 丙", False, "丙"))

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
                    except Exception:
                        time.sleep(1)
                        continue

                if not response or not response.text:
                    raise Exception("無法從 Google API 取得回應，請稍後再試。")
                
                data = json.loads(response.text)
                vehicles_list = data.get('vehicles', [])
                first_plate = vehicles_list[0].get('plate_no', 'Unknown').replace(' ', '_') if len(vehicles_list) > 0 else 'Unknown'
                
                date_std, date_zh, weekday_zh = parse_date_and_weekday(data.get('ta_date', ''))

                # 補全 T17A 車輛資訊
                label_mapping = {"甲": 0, "乙": 1, "丙": 2, "丁": 3, "戊": 4, "己": 5}
                for item in photo_items_for_t17:
                    if not item.get('is_scene'):
                        v_char = item.get('veh_label', '甲')
                        idx_num = label_mapping.get(v_char, 0)
                        if idx_num < len(vehicles_list):
                            v_obj = vehicles_list[idx_num]
                            item['veh_class_zh'] = v_obj.get('vehicle_class_zh', v_obj.get('vehicle_class_en', '車輛'))
                            item['plate_no'] = v_obj.get('plate_no', '')
                            item['damage'] = v_obj.get('damage', '車身毀爛')

                # 1. 生成 T17A Bytes
                t17_bytes = None
                if os.path.exists("T17_template.xlsx"):
                    t17_bytes = generate_t17_excel("T17_template.xlsx", data, photo_items_for_t17, app_mode)

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

                elif "TADO" in app_mode:
                    doc_obj = generate_tado_word(data, statement)
                    if doc_obj:
                        out_io = io.BytesIO()
                        doc_obj.save(out_io)
                        main_bytes = out_io.getvalue()
                        main_filename = f"TADO_附錄報告_{first_plate}.docx"
                        main_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                else: # T24
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

                            context[f"{{{{VEH_{i}_PLATE}}}}"] = v.get('plate_no', '———')
                            context[f"{{{{VEH_{i}_CLASS_ZH}}}}"] = v.get('vehicle_class_zh', v.get('vehicle_class_en', '———'))
                            context[f"{{{{VEH_{i}_CLASS_EN}}}}"] = v.get('vehicle_class_en', '———')
                            context[f"{{{{VEH_{i}_WEIGHT}}}}"] = v.get('gross_weight', '———')
                            context[f"{{{{VEH_{i}_MAKE}}}}"] = v.get('make', '———')
                            context[f"{{{{VEH_{i}_COLOR}}}}"] = v.get('color', '———')
                            context[f"{{{{VEH_{i}_TRANS_NO}}}}"] = v.get('transaction_no', '———')
                            context[f"{{{{VEH_{i}_EXPIRY}}}}"] = v.get('expiry_date', '———')
                            context[f"{{{{VEH_{i}_DAMAGE}}}}"] = v.get('damage', '———')
                            context[f"{{{{DRV_{i}_NAME_ZH}}}}"] = v.get('driver_name_zh', '———')
                            context[f"{{{{DRV_{i}_NAME_EN}}}}"] = v.get('driver_name_en', '———')
                            context[f"{{{{DRV_{i}_HKID}}}}"] = format_hkid_parentheses(v.get('hkid', '———'))
                            context[f"{{{{DRV_{i}_LIC}}}}"] = v.get('driving_license_no', '———')
                            context[f"{{{{DRV_{i}_ADDR_ZH}}}}"] = v.get('address_zh', '———')
                            context[f"{{{{DRV_{i}_ADDR_EN}}}}"] = v.get('address_en', '———')
                            context[f"{{{{DRV_{i}_PHONE}}}}"] = v.get('phone', '———')
                            context[f"{{{{DRV_{i}_STATEMENT}}}}"] = d_statement_formatted
                        else:
                            for key_suffix in ["PLATE", "CLASS_ZH", "CLASS_EN", "WEIGHT", "MAKE", "COLOR", "TRANS_NO", "EXPIRY", "DAMAGE"]:
                                context[f"{{{{VEH_{i}_{key_suffix}}}}"] = "———"
                            for key_suffix in ["NAME_ZH", "NAME_EN", "HKID", "LIC", "ADDR_ZH", "ADDR_EN", "PHONE", "STATEMENT"]:
                                context[f"{{{{DRV_{i}_{key_suffix}}}}"] = "———"

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

                # 3. 儲存所有資料至 Session State 供多次下載使用
                wa_text = f"🚓【交通意外報告已生成】\n• RN編號：{data.get('rn_no', '無')}\n• 意外時間：{data.get('ta_time', '無')}\n• 涉及車牌：{first_plate}\n• 模式：{app_mode.split(' ')[0]}\n請在附件接收報告檔。"
                
                st.session_state["main_bytes"] = main_bytes
                st.session_state["main_filename"] = main_filename
                st.session_state["main_mime"] = main_mime
                st.session_state["t17_bytes"] = t17_bytes
                st.session_state["t17_filename"] = f"IMSP-T17A_相片報告_{first_plate}.xlsx"
                st.session_state["t24_2_bytes"] = t24_2_bytes
                st.session_state["t24_2_filename"] = t24_2_filename
                st.session_state["wa_url"] = f"https://api.whatsapp.com/send?text={urllib.parse.quote(wa_text)}"
                st.session_state["preview_data"] = data
                st.session_state["generated"] = True
                
                st.rerun()

            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")

# ----------------- 4. 雙下載按鈕持久顯示區 -----------------
if st.session_state.get("generated"):
    st.subheader("📊 AI 自動提取與精準對齊預覽：")
    st.json(st.session_state["preview_data"])
    
    st.success("🎉 報告生成成功！請點擊下方按鈕下載（下載任何一份後頁面會保持，可繼續下載另一份）：")
    
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label=f"📥 1. 下載主報告 ({st.session_state['main_filename']})",
            data=st.session_state["main_bytes"],
            file_name=st.session_state["main_filename"],
            mime=st.session_state["main_mime"],
            key="dl_main_btn"
        )
    with col_dl2:
        if st.session_state["t17_bytes"]:
            st.download_button(
                label="🖼️ 2. 下載 IMSP-T17A 相片報告表 (XLSX)",
                data=st.session_state["t17_bytes"],
                file_name=st.session_state["t17_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_t17_btn"
            )
        else:
            st.warning("⚠️ 未能在 GitHub 找到 `T17_template.xlsx` 範本檔案。")

    if st.session_state.get("t24_2_bytes"):
        st.download_button(
            label=f"📥 1b. 下載 T24 正式報告續頁 ({st.session_state['t24_2_filename']})",
            data=st.session_state["t24_2_bytes"],
            file_name=st.session_state["t24_2_filename"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_t24_2_btn"
        )

    st.markdown(f'<a href="{st.session_state["wa_url"]}" target="_blank" class="wa-button">💬 分享案件通報至 WhatsApp 群組</a>', unsafe_allow_html=True)
    
    st.divider()
    if st.button("🔄 重新開始 / 處理下宗案件"):
        for k in state_keys:
            st.session_state[k] = state_keys[k]
        st.rerun()
