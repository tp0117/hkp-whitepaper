import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import os
import time
from PIL import Image

# 嘗試載入 python-docx 用於生成 Word 檔
try:
    import docx
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    docx = None

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
st.subheader("📸 2. 分類上載相片 (支援行車證、駕照、手填紀錄單、現場照片)")

tab_main, tab_veh_a, tab_veh_b, tab_veh_c = st.tabs([
    "📋 手填總表 / 現場紀錄照片",
    "🚗 車輛 甲 (Vehicle A) 相片",
    "🚙 車輛 乙 (Vehicle B) 相片",
    "🚚 其他 / 車輛 丙 (Vehicle C) 相片"
])

with tab_main:
    files_main = st.file_uploader("請上載【手填紀錄單 / 現場照片 / 總表】", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="main_files")

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
    doc = docx.Document()
    
    # 標題
    h = doc.add_heading('交通意外 (TADO) 附錄報告', level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    vehicles = data.get('vehicles', [])
    veh_labels = ["甲", "乙", "丙", "丁", "戊"]
    
    # 車輛及司機資料逐一生成
    for idx, v in enumerate(vehicles):
        label_char = veh_labels[idx] if idx < len(veh_labels) else str(idx + 1)
        doc.add_heading(f"車輛{label_char}：", level=2)
        
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.2
        
        # 1) 車牌 (含尾數)
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
        
        # 9) 駕駛執照號碼
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
        
        # 13) 司機口供
        d_statement = v.get('driver_statement', statement_text if statement_text else '———')
        p.add_run(f"13）司機口供： ").bold = True
        p.add_run(f"{d_statement}\n")

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

    images_with_tags.extend(process_files(files_main, "手填總表/現場照片"))
    images_with_tags.extend(process_files(files_veh_a, "車輛 甲 (Vehicle A)"))
    images_with_tags.extend(process_files(files_veh_b, "車輛 乙 (Vehicle B)"))
    images_with_tags.extend(process_files(files_veh_c, "車輛 丙 (Vehicle C)"))

    if not api_key:
        st.error("請先輸入或設定 Gemini API Key！")
    elif not images_with_tags and not statement:
        st.error("請最少上載一張相片或輸入文字資料！")
    else:
        with st.spinner('AI 正在全方位辨識（自動對齊車輛甲乙丙、規格與酒精測試紀錄）...'):
            try:
                client = genai.Client(api_key=api_key.strip())
                raw_images = [item[0] for item in images_with_tags]
                photo_notes = "\n".join([f"- 相片 {idx+1} 來自分類上載區：【{item[1]}】" for idx, item in enumerate(images_with_tags)])
                
                prompt = f"""
                你是一個專業的香港交通意外資料提取助手。
                請仔細分析所有上載的相片（包含行車證、駕駛執照、手填紙本表格/紀錄單、現場照片）以及現場文字內容：
                
                【上載相片分類備註】：
                {photo_notes}
                
                【核心交叉配對邏輯】：
                1. 請優先參考相片所在的分區標籤（如車輛甲、車輛乙）。
                2. 請精準對齊「車輛與司機」資料，車牌記得保留尾數字母 (例如 AB1234 A)。
                
                【補充文字內容與酒精測試】：
                {statement}
                
                【TADO 規格特別提取要求】：
                1. vehicle_class_en: 英文車輛類別 (如 PRIVATE CAR, LIGHT GOODS VEHICLE)。
                2. vehicle_class_zh: 對應中文車輛類別 (如 私家車, 輕型貨車)。
                3. gross_weight: 車輛總重量。若為私家車或電單車，請填寫 "———"；若為貨車，請填寫重量數字加單位 (如 "5.5T")。
                4. expiry_date: 行車證牌照期滿日期。若相片無顯示或未過期/無顯示，請填寫 "———"。
                5. breathalyzer_test: 警員是否有替涉及車輛司機進行酒精呼氣測試 (PBT) 的紀錄或結果。若有提及請詳細填寫，若無請填寫 "未有提供 / 無紀錄"。
                6. police_veh_no: 警車車牌 (例如 AM 8810)。
                7. police_no: 警員編號 (例如 警員26088)。
                
                【地址翻譯要求】：
                - 英文地址必須翻譯成香港繁體中文地址填入 address_zh。
                - 中文地址必須翻譯成英文地址填入 address_en。
                
                請嚴格以 JSON 格式輸出：
                {{
                    "log_no": "Log No 紀錄號碼 (無則留空)",
                    "rn_no": "HKP RN 報案號碼/案號 (如 RN:26031028)",
                    "ta_date": "意外日期",
                    "ta_time": "意外時間",
                    "location": "意外地點",
                    "police_veh_no": "警車車牌 (如 AM 8810)",
                    "police_no": "到場警員編號 (如 警員26088)",
                    "breathalyzer_test": "酒精呼氣測試紀錄/結果",
                    "vehicles": [
                        {{
                            "plate_no": "車牌號碼 (包含尾數，如 AB1234 A)",
                            "vehicle_class_en": "英文車輛類別 (如 PRIVATE CAR)",
                            "vehicle_class_zh": "中文車輛類別 (如 私家車)",
                            "gross_weight": "總重量 (私家車填 ———，貨車填 5.5T)",
                            "make": "車輛廠名 (如 HONDA)",
                            "color": "車輛顏色 (如 SILVER/VAR)",
                            "transaction_no": "存檔編號 Transaction Number",
                            "expiry_date": "屆滿日期 (無則填 ———)",
                            "damage": "車輛損毀情況",
                            "driver_name_en": "司機英文姓名",
                            "driver_name_zh": "司機中文姓名",
                            "phone": "電話號碼",
                            "hkid": "身份證號碼",
                            "driving_license_no": "駕駛執照號碼 (如 Y463269 7)",
                            "address_zh": "完整中文住址",
                            "address_en": "完整英文住址",
                            "driver_statement": "司機口供摘要"
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
                
                st.subheader("📊 AI 自動提取結果預覽：")
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
                    
                    if data.get('log_no'): ws_data['B1'] = data.get('log_no')
                    if data.get('rn_no'): ws_data['B2'] = data.get('rn_no')
                    if data.get('ta_date'): ws_data['B3'] = data.get('ta_date')
                    if data.get('ta_time'): ws_data['B4'] = data.get('ta_time')
                    if data.get('location'): ws_data['B5'] = data.get('location')
                    
                    if data.get('police_no'): ws_data['E3'] = data.get('police_no')
                    
                    v_mappings = [
                        {"plate": "B12", "veh_class": "B13", "make": "B14", "color": "B15", "trans_no": "B16", "expiry": "B17", "seating": "B18", "damage": "B20", "name_en": "E12", "name_zh": "E13", "age": "E14", "addr_zh": "E15", "addr_en": "E16", "phone": "E17", "hkid": "E18", "lic_no": "E19", "dl_class": "E20"},
                        {"plate": "B27", "veh_class": "B28", "make": "B29", "color": "B30", "trans_no": "B31", "expiry": "B32", "seating": "B33", "damage": "B35", "name_en": "E27", "name_zh": "E28", "age": "E29", "addr_zh": "E30", "addr_en": "E31", "phone": "E32", "hkid": "E33", "lic_no": "E34", "dl_class": "E35"},
                        {"plate": "B43", "veh_class": "B44", "make": "B45", "color": "B46", "trans_no": "B47", "expiry": "B48", "seating": "B49", "damage": "B51", "name_en": "E43", "name_zh": "E44", "age": "E45", "addr_zh": "E46", "addr_en": "E47", "phone": "E48", "hkid": "E49", "lic_no": "E50", "dl_class": "E51"}
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
                        ws_data[m["damage"]] = v.get('damage', '')
                        
                        ws_data[m["name_en"]] = v.get('driver_name_en', '')
                        ws_data[m["name_zh"]] = v.get('driver_name_zh', '')
                        ws_data[m["addr_zh"]] = v.get('address_zh', '')
                        ws_data[m["addr_en"]] = v.get('address_en', '')
                        ws_data[m["phone"]] = v.get('phone', '')
                        if v.get('hkid'): ws_data[m["hkid"]] = v.get('hkid')
                        ws_data[m["lic_no"]] = v.get('driving_license_no', '')

                    if statement:
                        ws_police.cell(row=5, column=1, value=f"現場紀錄及司機口供：\n{statement}")
                    
                    wb.save(output_path)
                    st.success(f"🎉 TAPI 白皮書生成成功！(模型：{used_model_name})")
                    
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
