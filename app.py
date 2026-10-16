# ----------------- UI 手機化美化 (Custom CSS) -----------------
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
    
    /* 放大生成按鈕，適合手機大拇指點擊 */
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
    
    /* 強化單選框與 Tab 的手機點擊範圍 */
    div[role="radiogroup"] {
        padding: 8px;
        background-color: #f0f2f6;
        border-radius: 8px;
    }
    </style>
""", unsafe_allow_html=True)
