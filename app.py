import streamlit as st
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from google import genai
from google.genai import types
import json
import os
import time
import re
from PIL import Image

# 嘗試載入 python-docx 用於生成 Word 檔
try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except Exception:
    docx = None

# Helper 函數：僅用於身份證號碼尾數加上圓括號，如 G111759(8)
def format_hkid_parentheses(code_str):
    if not code_str:
        return ""
    s = str(code_str).strip()
    # 若已經有尾數括號 (X) 則直接傳回
    if re.search(r'\([0-9A-Za-z]\)$', s):
        return s
    # 比對前面 7-8 位號碼 + 最後一位校驗碼 (如 G111759 8, G1117598, Y463269-7)
    match = re.match(r'^([A-Za-z]{1,2}\d{6})[\s\-_]?([0-9A-Za-z])$', s)
    if match:
        return f"{match.group(1)}({match.group(2)})"
    return s

# 網頁標題及頁面設定
st.set_page_config(page_title="交通意外資料處理系統 (TAPI/TADO)", page_icon="🚓", layout="wide")

# 🔒 密碼鎖設定：優先從 Secrets 讀取 APP_PASSWORD，若未設定則預設為 "8888"
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "8888")

# 初始化登入狀態
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# 若未通過驗證，顯示密碼輸入畫面
if not st.session_state["authenticated"]:
    st.title("🔒 白皮書 / TADO 系統 - 安全驗證")
    st.info("本系統僅供內部授權人員使用，請輸入訪問密碼。")
    
    input_pwd = st.text_input("請輸入訪問密碼：", type="password")
    if st.button("🔑 登入"):
        if input_pwd == APP_PASSWORD:
            st.session_state["authenticated"] = True
            st.success("驗證成功！正在載入系統...")
            st.rerun()
        else:
            st.error("密碼錯誤，請重新輸入！")
    
    st.stop()  # 密碼未通過前，停止執行後續所有系統邏輯

# ----------------- 以下為驗證通過後的系統主介面 -----------------

st.title("🚓 交通意外資料處理系統 (TAPI 白皮書 / TADO Word 附錄)")

# 優先從 Secrets 讀取 Gemini API Key
default_api_key = st.secrets.get("GEMINI_API_KEY", "")

# 左側欄設定
with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input("Gemini API Key", value=default_api_key, type="password", help="系統已自動讀取雲端預設 Key，亦可手動覆蓋")
    st.markdown("[按此免費獲取 Google Gemini API Key](https://aistudio.google.com/)")
    
    st.divider()
    
    # 🤖 允許使用者自選 Gemini AI 模型
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
    ["TAPI (有人受傷 - 生成 Excel 白皮書)", "TADO (僅車損 - 生成 Word 附錄)"],
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

# Helper: 生成精準排版的 TADO Word 附錄檔
def generate_tado_word(data, statement_text):
    if not docx:
        return None
    try:
        doc = docx.Document()
        
        # 標題
        h = doc.add_heading('交通意外 (TADO) 附錄報告', level=1)
        if hasattr(WD_ALIGN_PARAGRAPH, 'CENTER'):
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        vehicles = data.get('vehicles', [])
        veh_labels = ["甲", "乙", "丙", "丁", "戊"]
        
        # 車輛及司機資料逐一生成
        for idx, v in enumerate(vehicles):
            label_char = veh_labels[idx] if idx < len(veh_labels) else str(idx + 1)
            doc.add_heading(f"車輛{label_char}：", level=2)
            
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = 1.2
            
            # 1) 車牌 (必須包含尾數)
            plate = v.get('plate_no', '———')
            p.add_run(f"1）車牌： ").bold = True
            p.add_run(f"{plate}\n")
            
            # 2) 車輛類型 (英/中)
            v_class_en = v.get('vehicle_class_en', v.get('vehicle_class', '———'))
            v_class_zh = v.get('vehicle_class_zh', '———')
            v_type_str = f"{v_class_en} / {v_class_zh}" if v_class_zh != '———' else v_class_en
            p.add_run(f"2）車輛類型： ").bold = True
            p.add_run(f"{v_type_str}\n")
            
            # 3) 總重量 (私家車/電單車為 ———，貨車填重量如 5.5T)
            weight = v.get('gross_weight', '———')
            p.add_run(f"3）總重量： ").bold = True
            p.add_run(f"{weight}\n")
            
            # 4) 廠名
            make = v.get('make', '———')
            p.add_run(f"4）廠名： ").bold = True
            p.add_run(f"{make}\n")
            
            # 5) 顏色
            color = v.get('color', '———')
            p.add_run(f"5）顏色： ").bold = True
            p.add_run(f"{color}\n")
            
            # 6) 存檔編號 (Transaction No.)
            trans_no = v.get('transaction_no', '———')
            p.add_run(f"6）存檔編號： ").bold = True
            p.add_run(f"{trans_no}\n")
            
            # 7) 屆滿日期 (無則 ———)
            exp = v.get('expiry_date', '———')
            p.add_run(f"7）屆滿日期： ").bold = True
            p.add_run(f"{exp}\n")
            
            # 8) 駕駛人 (英文名及中文名)
            d_en = v.get('driver_name_en', '')
            d_zh = v.get('driver_name_zh', '')
            driver_str = f"{d_en} ({d_zh})" if d_zh else (d_en if d_en else '———')
            p.add_run(f"8）駕駛人： ").bold = True
            p.add_run(f"{driver_str}\n")
            
            # 9) 駕駛執照號碼 (保持原樣，不加尾數括號)
            lic_no = v.get('driving_license_no', '———')
            p.add_run(f"9）駕駛執照號碼： ").bold = True
            p.add_run(f"{lic_no}\n")
            
            # 10) 地址 (中英文雙向翻譯)
            addr_zh = v.get('address_zh', '———')
            addr_en = v.get('address_en', '———')
            p.add_run(f"10）地址： ").bold = True
            p.add_run(f"\n   中文：{addr_zh}\n   英文：{addr_en}\n")
            
            # 11) 電話號碼
            phone = v.get('phone', '———')
            p.add_run(f"11）電話號碼： ").bold = True
            p.add_run(f"{phone}\n")
            
            # 12) 車輛損毀情況
            damage = v.get('damage', '———')
            p.add_run(f"12）車輛損毀情況： ").bold = True
            p.add_run(f"{damage}\n")
            
            # 13) 司機口供 (自動以「」括號括住)
            d_statement = v.get('driver_statement', statement_text if statement_text else '———')
            if d_statement and d_statement != '———':
                s_clean = d_statement.strip()
                if not (s_clean.startswith('「') and s_clean.endswith('」')):
                    d_statement_formatted = f"「{s_clean}」"
                else:
                    d_statement_formatted = s_clean
            else:
                d_statement_formatted = '———'
                
            p.add_run(f"13）司機口供： ").bold = True
            p.add_run(f"{d_statement_formatted}\n")

        # 結尾：警方資料及酒精呼氣測試
        doc.add_heading("警方資料及測試紀錄：", level=2)
        p_police = doc.add_paragraph()
        p_police.paragraph_format.line_spacing = 1.2
        
        police_veh = data.get('police_veh_no', '———')
        police_officer = data.get('police_no', '———')
        rn_no = data.get('rn_no', '———')
        pbt = data.get('breathalyzer_test', '未有提供 / 無紀錄')
        
        p_police.add_run(f"• 警車車牌： ").bold = True
        p_police.add_run(f"{police_veh}\n")
        p_police.add_run(f"• 到場警員： ").bold = True
        p_police.add_run(f"{police_officer}\n")
        p_police.add_run(f"• 報案號碼 (RN)： ").bold = True
        p_police.add_run(f"{rn_no}\n")
        p_police.add_run(f"• 酒精呼氣測試 (PBT)： ").bold = True
        p_police.add_run(f"{pbt}\n")

        return doc
    except Exception as err:
        st.error(f"建立 Word 文件時發生錯誤：{str(err)}")
        return None

# 生成按鈕
button_label = "🚀 立即生成 Excel 白皮書" if "TAPI" in app_mode else "📄 立即生成 Word 附錄報告"

if st.button(button_label):
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
                   - escort: 護送人/遞送人/身份備註 (如 M/C WW5031 司機)
                
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
                
                candidate_models = [selected_model, "gemini-1.5-flash", "gemini-2.5-pro"]
                candidate_models = list(dict.fromkeys(candidate_models))
                
                response = None
                used_model_name = None
                last_exception = None

                for m_name in candidate_models:
                    try:
                        res = client.models.generate_content(
                            model=m_name,
                            contents=contents,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json"
                            )
                        )
                        if res and res.text:
                            response = res
                            used_model_name = m_name
                            break
                    except Exception as e:
                        last_exception = e
                        time.sleep(1)
                        continue

                if not response or not response.text:
                    raise last_exception if last_exception else Exception("無法從 Google API 取得回應，請稍後再試。")
                
                data = json.loads(response.text)
                
                st.subheader("📊 AI 自動提取與精準對齊預覽：")
                st.json(data)
                
                vehicles_list = data.get('vehicles', [])
                first_plate = vehicles_list[0].get('plate_no', 'Unknown').replace(' ', '_') if len(vehicles_list) > 0 else 'Unknown'
                
                # 分流處理：TAPI (Excel 白皮書) 或 TADO (Word 附錄)
                if "TAPI" in app_mode:
                    template_path = "T26_白皮書_template_v2.xlsx"
                    output_path = f"白皮書_{first_plate}.xlsx"
                    
                    wb = openpyxl.load_workbook(template_path)
                    ws_data = wb['Data']
                    ws_police = wb['Police']
                    
                    # 1. 頂部案件與隧道時間
                    if data.get('log_no'): ws_data['B1'] = data.get('log_no')
                    if data.get('rn_no'): ws_data['B2'] = data.get('rn_no')
                    if data.get('ta_date'): ws_data['B3'] = data.get('ta_date')
                    if data.get('ta_time'): ws_data['B4'] = data.get('ta_time')
                    if data.get('location'): ws_data['B5'] = data.get('location')
                    if data.get('tunnel_staff_arrived'): ws_data['B6'] = data.get('tunnel_staff_arrived')
                    if data.get('tunnel_staff_departed'): ws_data['B7'] = data.get('tunnel_staff_departed')
                    
                    # 2. 警察資料精準獨立拆解 (E 欄)
                    if data.get('police_arrived_time'): ws_data['E2'] = data.get('police_arrived_time')
                    if data.get('police_no'): ws_data['E3'] = data.get('police_no')          # 普通警員
                    if data.get('ai_staff_no'): ws_data['E4'] = data.get('ai_staff_no')        # AI Staff
                    if data.get('ai_arrived_time'): ws_data['E5'] = data.get('ai_arrived_time')
                    if data.get('ai_departed_time'): ws_data['E6'] = data.get('ai_departed_time')
                    
                    # 3. 車輛 1 ~ 4 儲存格精準映射 (Driver 1: E20/E21, Driver 2: E35/E36, Driver 3: E51/E52, Driver 4: E66/E67)
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
                        
                        # 身份證號碼：帶尾數括號；駕駛執照號碼：保持原樣不帶括號
                        hkid_val = format_hkid_parentheses(v.get('hkid', ''))
                        lic_val = v.get('driving_license_no', '')
                        
                        if hkid_val: ws_data[m["hkid"]] = hkid_val
                        if lic_val: ws_data[m["lic_no"]] = lic_val
                        if v.get('dl_class'): ws_data[m["dl_class"]] = v.get('dl_class')
                        if v.get('dl_expiry_date'): ws_data[m["dl_exp"]] = v.get('dl_expiry_date')

                    # 4. 傷者資料 100% 精準對齊 (含電話號碼與雙向翻譯地址)
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
                        ws_data[im["ph"]] = inj.get('phone', '')                # 傷者電話號碼
                        ws_data[im["vno"]] = inj.get('vehicle_no', '')
                        ws_data[im["hosp"]] = inj.get('hospital', '')
                        ws_data[im["esc"]] = inj.get('escort', '')

                    if statement:
                        ws_police.cell(row=5, column=1, value=f"現場紀錄及司機口供：\n{statement}")
                    
                    # 5. 精準將 car_diagram.png 車仔圖形分別根據獨立指定尺寸植入 P.1 (I41) 及 P.2 (H14, H39)
                    if os.path.exists("car_diagram.png"):
                        if 'P.1' in wb.sheetnames:
                            ws_p1 = wb['P.1']
                            img_p1 = OpenpyxlImage("car_diagram.png")
                            img_p1.width = 414   # 10.95 公分
                            img_p1.height = 347  # 9.17 公分
                            ws_p1.add_image(img_p1, 'I41')
                            
                        if 'P.2' in wb.sheetnames:
                            ws_p2 = wb['P.2']
                            
                            # P.2 車輛 2 (H14)
                            img_p2_1 = OpenpyxlImage("car_diagram.png")
                            img_p2_1.width = 434   # 11.49 公分
                            img_p2_1.height = 355  # 9.39 公分
                            ws_p2.add_image(img_p2_1, 'H14')
                            
                            # P.2 車輛 3 (H39)
                            img_p2_2 = OpenpyxlImage("car_diagram.png")
                            img_p2_2.width = 437   # 11.57 公分
                            img_p2_2.height = 334  # 8.84 公分
                            ws_p2.add_image(img_p2_2, 'H39')
                    
                    wb.save(output_path)
                    st.success(f"🎉 TAPI 白皮書生成成功！(已完美對齊 IMSP-T3 頂部車輛類別與傷者資訊！)")
                    
                    with open(output_path, "rb") as file:
                        st.download_button(
                            label="📥 下載全功能 Excel 白皮書",
                            data=file,
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    if os.path.exists(output_path):
                        os.remove(output_path)

                else:
                    # 生成 TADO Word 附錄
                    doc_obj = generate_tado_word(data, statement)
                    if doc_obj is None:
                        st.error("⚠️ 無法建立 Word 檔案！請確定 GitHub 的 requirements.txt 已加入 python-docx 套件，並且 Streamlit 已完成重新部署。")
                    else:
                        output_word_path = f"TADO_附錄報告_{first_plate}.docx"
                        doc_obj.save(output_word_path)
                        
                        st.success(f"🎉 TADO Word 附錄生成成功！(模型：{used_model_name})")
                        
                        with open(output_word_path, "rb") as file:
                            st.download_button(
                                label="📥 下載 Word 附錄報告 (.docx)",
                                data=file,
                                file_name=output_word_path,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        if os.path.exists(output_word_path):
                            os.remove(output_word_path)
                
            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")
