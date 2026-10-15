import streamlit as st
import openpyxl
from google import genai
from google.genai import types
import json
import os
from PIL import Image

# 網頁標題及頁面設定
st.set_page_config(page_title="白皮書自動生成系統", page_icon="📝", layout="wide")

# 🔒 密碼鎖設定：優先從 Secrets 讀取 APP_PASSWORD，若未設定則預設為 "8888"
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "8888")

# 初始化登入狀態
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

# 若未通過驗證，顯示密碼輸入畫面
if not st.session_state["authenticated"]:
    st.title("🔒 白皮書系統 - 安全驗證")
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

st.title("🚓 交通意外白皮書自動生成系統 (支援手填表格相片)")

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
        "🤖 選擇 AI 模型",
        ["gemini-2.5-flash", "gemini-3.8-flash", "gemini-2.5-pro"],
        index=0,
        help="若目前模型跳出額度滿或過載提示，可在這裡切換至其他模型續用"
    )
    
    st.divider()
    if st.button("🔒 登出系統"):
        st.session_state["authenticated"] = False
        st.rerun()

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
        with st.spinner(f'AI ({selected_model}) 正在全方位辨識相片（含行車證詳細規格、駕照 Full Class 類別及手寫字）...'):
            try:
                client = genai.Client(api_key=api_key.strip())
                images = [Image.open(file) for file in uploaded_files] if uploaded_files else []
                
                prompt = f"""
                你是一個專業的香港交通意外白皮書資料提取助手。
                請仔細分析所有上載的相片（特別是行車證、駕駛執照、手填紙本表格/紀錄單、現場照片）以及現場文字內容：
                
                【補充文字內容】：
                {statement}
                
                【行車證 (Vehicle Licence) 特別提取要求】：
                1. plate_no：車牌號碼 (如 UG8125)。
                2. vehicle_class：行車證上的 Class / 車輛類別 (如 PRIVATE CAR, LIGHT GOODS VEHICLE, TAXI 等)。
                3. make：車輛廠名 (如 TOYOTA, HYUNDAI)。
                4. color：車輛顏色 (如 SILVER)。
                5. transaction_no：行車證上的車輛牌照號碼 / 交易編號 Transaction No. (如 C06-0123)。
                6. expiry_date：行車證牌照期滿日期 (如 DD/MM/YYYY 或 DD-MM-YYYY)。
                7. seating_capacity：行車證上的車輛載客額 Seating Capacity (如數字 5 或 4)。
                
                【駕駛執照 (Driving Licence) 特別提取要求】：
                1. driving_license_no：駕駛執照號碼 (請移除括號)。
                2. dl_class：駕駛執照上「正式執照 Full class」欄位內的數字類別 (如 1, 2 或 1, 2, 18 等)。
                
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
                            "vehicle_class": "行車證車輛類別 Class (如 LIGHT GOODS VEHICLE)",
                            "make": "車輛廠名 (如 TOYOTA)",
                            "color": "車輛顏色 (如 SILVER)",
                            "transaction_no": "車輛牌照號碼 / Transaction No.",
                            "expiry_date": "期滿日期",
                            "seating_capacity": "車輛載客額 Seating Capacity",
                            "damage": "該車損毀性質及程度",
                            "driver_name_en": "司機英文姓名",
                            "driver_name_zh": "司機中文姓名",
                            "age": "年齡 (無則留空)",
                            "phone": "電話號碼",
                            "hkid": "身份證號碼",
                            "driving_license_no": "駕駛執照號碼 (請移除括號)",
                            "dl_class": "駕駛執照 Full Class 數字類別 (如 1, 2, 18)",
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
                
                # 直接調用使用者在選單中選取的模型
                response = client.models.generate_content(
                    model=selected_model,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    )
                )
                
                if not response or not response.text:
                    st.error("AI 未能傳回有效的資料，請重試或切換模型。")
                    st.stop()
                
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
                
                # 2. 車輛 1 ~ 4 儲存格映射表 (完全對齊 Excel A 欄標題)
                v_mappings = [
                    {"plate": "B12", "veh_class": "B13", "make": "B14", "color": "B15", "trans_no": "B16", "expiry": "B17", "seating": "B18", "damage": "B20", "name_en": "E12", "name_zh": "E13", "age": "E14", "addr_zh": "E15", "addr_en": "E16", "phone": "E17", "hkid": "E18", "lic_no": "E19", "dl_class": "E20"},
                    {"plate": "B27", "veh_class": "B28", "make": "B29", "color": "B30", "trans_no": "B31", "expiry": "B32", "seating": "B33", "damage": "B35", "name_en": "E27", "name_zh": "E28", "age": "E29", "addr_zh": "E30", "addr_en": "E31", "phone": "E32", "hkid": "E33", "lic_no": "E34", "dl_class": "E35"},
                    {"plate": "B43", "veh_class": "B44", "make": "B45", "color": "B46", "trans_no": "B47", "expiry": "B48", "seating": "B49", "damage": "B51", "name_en": "E43", "name_zh": "E44", "age": "E45", "addr_zh": "E46", "addr_en": "E47", "phone": "E48", "hkid": "E49", "lic_no": "E50", "dl_class": "E51"},
                    {"plate": "B58", "veh_class": "B59", "make": "B60", "color": "B61", "trans_no": "B62", "expiry": "B63", "seating": "B64", "damage": "B66", "name_en": "E58", "name_zh": "E59", "age": "E60", "addr_zh": "E61", "addr_en": "E62", "phone": "E63", "hkid": "E64", "lic_no": "E65", "dl_class": "E66"}
                ]
                
                for idx, v in enumerate(vehicles_list):
                    if idx >= len(v_mappings): break
                    m = v_mappings[idx]
                    ws_data[m["plate"]] = v.get('plate_no', '')
                    if v.get('vehicle_class'): ws_data[m["veh_class"]] = v.get('vehicle_class')
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
                    if v.get('hkid'): ws_data[m["hkid"]] = v.get('hkid')
                    ws_data[m["lic_no"]] = v.get('driving_license_no', '')
                    if v.get('dl_class'): ws_data[m["dl_class"]] = v.get('dl_class')

                # 3. 傷者資料填寫 (傷者 1~4)
                if data.get('injured_total'): ws_data['B74'] = data.get('injured_total')
                
                injured_list = data.get('injured_list', [])
                i_mappings = [
                    {"en": "B77", "zh": "B78", "age": "B79", "azh": "B80", "aen": "B81", "ph": "B82", "vno": "B83", "hosp": "B84", "esc": "B85"},
                    {"en": "E77", "zh": "E78", "age": "E79", "azh": "E78", "aen": "E81", "ph": "E82", "vno": "E83", "hosp": "E84", "esc": "E85"},
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
                
                st.success(f"🎉 白皮書生成成功！(使用的是 {selected_model} 模型)")
                
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
                err_str = str(e)
                if "429" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    st.error(f"⚠️ 模型 {selected_model} 今日免費次數已滿（429）。請在左側選單切換至其他模型（如 gemini-2.5-flash 或 gemini-2.5-pro）再按一次生成！")
                elif "503" in err_str or "UNAVAILABLE" in err_str:
                    st.error(f"⚠️ Google 伺服器短暫過載（503）。請等 5 秒後重試，或在左側選單切換其他模型。")
                else:
                    st.error(f"發生錯誤：{err_str}")
