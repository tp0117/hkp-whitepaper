import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import os
from PIL import Image

# 網頁標題及設定
st.set_page_config(page_title="白皮書自動生成系統", page_icon="📝", layout="wide")
st.title("🚓 交通意外白皮書自動生成系統 (支援手填表格相片)")

# 優先從 Secrets 讀取 API Key
default_api_key = st.secrets.get("GEMINI_API_KEY", "")

# 左側欄設定 API Key
with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input("Gemini API Key", value=default_api_key, type="password", help="系統已自動讀取雲端預設 Key，亦可手動覆蓋")
    st.markdown("[按此免費獲取 Google Gemini API Key](https://aistudio.google.com/)")

# 準備介面元件
st.subheader("1. 上載相片 (含行車證、駕照、手填紀錄表或現場照片)")
uploaded_files = st.file_uploader("請上載所有相關相片 (可一次選取多張，包含手填紙本表格相片)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

st.subheader("2. 司機口供及補充文字 (可選填)")
statement = st.text_area("請輸入額外口供或補充資料 (若相片已包含所有資料，此處可留空)", height=120)

# 生成按鈕
if st.button("🚀 立即生成 Excel 白皮書"):
    if not api_key:
        st.error("請先輸入或設定 Gemini API Key！")
    elif not uploaded_files and not statement:
        st.error("請上載最少一張相片或輸入文字資料！")
    else:
        with st.spinner('AI 正在全方位辨識相片（包含手填表格手寫字）、提取時間地點並翻譯地址中...'):
            try:
                client = genai.Client(api_key=api_key.strip())
                images = [Image.open(file) for file in uploaded_files] if uploaded_files else []
                
                prompt = f"""
                你是一個專業的香港交通意外白皮書資料提取助手。
                請仔細分析所有上載的相片（包含行車證、駕駛執照、手填紙本表格/紀錄單、現場照片）以及現場文字內容：
                
                【補充文字內容】：
                {statement}
                
                【特別說明】：
                相片中可能包含「手填/手寫的官方或公司紀錄表」，請精準讀取上面的手寫字體與數字，提取以下所有欄位資訊。
                
                【警察及時間資料要求】：
                1. police_no（普通警員編號）：只提取軍裝/普通警員編號 (如 SPC 48840, PC 26909)。【絕對不可包含 AM 開頭警車車牌】。
                2. ai_staff_no（AI 人員編號）：交通意外調查組人員編號。
                3. 各項時間與地點欄位請照實提取。
                
                【地址翻譯要求】：
                - 英文地址必須翻譯成香港繁體中文地址填入 address_zh。
                - 中文地址必須翻譯成英文地址填入 address_en。
                
                請嚴格以 JSON 格式輸出：
                {{
                    "log_no": "Log No 紀錄號碼 (無則留空)",
                    "rn_no": "HKP RN 報案號碼/案號 (如 26031881)",
                    "ta_date": "TA Date 意外日期 (如 YYYY-MM-DD 或 DD/MM/YYYY)",
                    "ta_time": "TA Time 意外時間 (如 HH:MM)",
                    "location": "Location 意外地點",
                    "tunnel_staff_arrived": "隧道職員到達時間",
                    "tunnel_staff_departed": "隧道職員離開時間",
                    "police_arrived_time": "Police Arrived Time 警察到達時間",
                    "police_no": "普通警員編號 (切勿包含 AM 車牌)",
                    "ai_staff_no": "AI Staff No 交通意外調查組人員編號",
                    "ai_arrived_time": "HKP AI 到達時間",
                    "ai_departed_time": "HKP AI 離開時間",
                    "vehicles": [
                        {{
                            "plate_no": "車牌號碼",
                            "make": "車輛廠名 (如 TOYOTA)",
                            "color": "車輛顏色",
                            "damage": "該車損毀性質及程度",
                            "driver_name_en": "司機英文姓名",
                            "driver_name_zh": "司機中文姓名",
                            "age": "年齡 (無則留空)",
                            "phone": "電話號碼",
                            "hkid": "身份證號碼",
                            "driving_license_no": "駕駛執照號碼 (請移除括號)",
                            "address_zh": "完整中文住址 (繁體中文)",
                            "address_en": "完整英文住址"
                        }}
                    ],
                    "injured_total": "受傷者合共人數 (無人受傷填 0)",
                    "injured_list": [
                        {{
                            "name_en": "傷者英文姓名",
                            "name_zh": "傷者中文姓名",
                            "age": "年齡",
                            "address_zh": "中文住址",
                            "address_en": "英文住址",
                            "phone": "電話號碼",
                            "vehicle_no": "涉及車輛號碼/車牌",
                            "hospital": "送往 Hospital 醫院",
                            "escort": "遞送人/護送人"
                        }}
                    ]
                }}
                """
                
                contents = images + [prompt]
                
                response = client.models.generate_content(
                    model='gemini-2.0-flash',
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    )
                )
                
                data = json.loads(response.text)
                
                st.subheader("📊 AI 提取與手寫辨識結果預覽：")
                st.json(data)
                
                template_path = "T26_白皮書_template_v2.xlsx"
                
                vehicles_list = data.get('vehicles', [])
                first_plate = vehicles_list[0].get('plate_no', 'Unknown') if len(vehicles_list) > 0 else 'Unknown'
                output_path = f"白皮書_{first_plate}.xlsx"
                
                wb = openpyxl.load_workbook(template_path)
                ws_data = wb['Data']
                ws_police = wb['Police']
                
                # 1. 頂部案件、警察及各項到達/離開時間資訊
                if data.get('log_no'): ws_data['B1'] = data.get('log_no')
                if data.get('rn_no'): ws_data['B2'] = data.get('rn_no')
                if data.get('ta_date'): ws_data['B3'] = data.get('ta_date')
                if data.get('ta_time'): ws_data['B4'] = data.get('ta_time')
                if data.get('location'): ws_data['B5'] = data.get('location')
                if data.get('tunnel_staff_arrived'): ws_data['B6'] = data.get('tunnel_staff_arrived')
                if data.get('tunnel_staff_departed'): ws_data['B7'] = data.get('tunnel_staff_departed')
                
                if data.get('police_arrived_time'): ws_data['E2'] = data.get('police_arrived_time')
                if data.get('police_no'): ws_data['E3'] = data.get('police_no')
                if data.get('ai_staff_no'): ws_data['E4'] = data.get('ai_staff_no')
                if data.get('ai_arrived_time'): ws_data['E5'] = data.get('ai_arrived_time')
                if data.get('ai_departed_time'): ws_data['E6'] = data.get('ai_departed_time')
                
                # 2. 車輛 1 ~ 4 儲存格映射表
                v_mappings = [
                    {"plate": "B12", "make": "B14", "color": "B15", "damage": "B20", "name_en": "E12", "name_zh": "E13", "age": "E14", "addr_zh": "E15", "addr_en": "E16", "phone": "E17", "hkid": "E18", "lic_no": "E19"},
                    {"plate": "B27", "make": "B29", "color": "B30", "damage": "B35", "name_en": "E27", "name_zh": "E28", "age": "E29", "addr_zh": "E30", "addr_en": "E31", "phone": "E32", "hkid": "E33", "lic_no": "E34"},
                    {"plate": "B43", "make": "B45", "color": "B46", "damage": "B51", "name_en": "E43", "name_zh": "E44", "age": "E45", "addr_zh": "E46", "addr_en": "E47", "phone": "E48", "hkid": "E49", "lic_no": "E50"},
                    {"plate": "B58", "make": "B60", "color": "B61", "damage": "B66", "name_en": "E58", "name_zh": "E59", "age": "E60", "addr_zh": "E61", "addr_en": "E62", "phone": "E63", "hkid": "E64", "lic_no": "E65"}
                ]
                
                for idx, v in enumerate(vehicles_list):
                    if idx >= len(v_mappings): break
                    m = v_mappings[idx]
                    ws_data[m["plate"]] = v.get('plate_no', '')
                    ws_data[m["make"]] = v.get('make', '')
                    ws_data[m["color"]] = v.get('color', '')
                    ws_data[m["damage"]] = v.get('damage', '')
                    ws_data[m["name_en"]] = v.get('driver_name_en', '')
                    ws_data[m["name_zh"]] = v.get('driver_name_zh', '')
                    if v.get('age'): ws_data[m["age"]] = v.get('age')
                    ws_data[m["addr_zh"]] = v.get('address_zh', '')
                    ws_data[m["addr_en"]] = v.get('address_en', '')
                    ws_data[m["phone"]] = v.get('phone', '')
                    if v.get('hkid'): ws_data[m["hkid"]] = v.get('hkid')
                    ws_data[m["lic_no"]] = v.get('driving_license_no', '')

                # 3. 傷者資料填寫 (傷者 1~4)
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

                # 4. 填入口供頁 (Police 頁面)
                if statement:
                    ws_police.cell(row=5, column=1, value=f"現場紀錄及司機口供：\n{statement}")
                
                wb.save(output_path)
                
                st.success("🎉 白皮書生成成功！手填表格相片上的時間與案件資料已自動精準填入！")
                
                # 下載按鈕
                with open(output_path, "rb") as file:
                    st.download_button(
                        label="📥 下載全功能 Excel 白皮書",
                        data=file,
                        file_name=output_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                if os.path.exists(output_path):
                    os.remove(output_path)
                
            except Exception as e:
                st.error(f"發生錯誤：{str(e)}")
