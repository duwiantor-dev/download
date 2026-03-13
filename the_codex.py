import streamlit as st

st.set_page_config(page_title="CODEX.ID", layout="wide")

CODEX_URL = "https://drive.google.com/uc?export=download&id=1Kvrc6OlMMk0RIyuXrfiD2V-3HgdP-Lgh"
ADDON_URL = "https://docs.google.com/spreadsheets/d/1yMGSboto7eVryyx6m1QlgHDNHSH3mj8r/export?format=xlsx"

st.markdown(
    """
    <style>

        /* HILANGKAN HEADER PUTIH STREAMLIT */
        header[data-testid="stHeader"]{
            display:none;
        }

        .stApp{
            background-color:#f7f7f8;
        }

        .block-container{
            padding-top:0rem;
            padding-bottom:2rem;
            max-width:1100px;
        }

        .main-title{
            font-size:2.8rem;
            font-weight:800;
            color:#1f2a44;
            margin-top:10px;
            margin-bottom:4px;
            line-height:1.1;
        }

        .subtitle{
            color:#8b8f97;
            font-size:1rem;
            margin-bottom:20px;
        }

        .info-box{
            background:white;
            border:1px solid #d9dde3;
            border-radius:10px;
            overflow:hidden;
            margin-bottom:20px;
        }

        .info-head{
            background:#f3f4f6;
            border-bottom:1px solid #d9dde3;
            padding:12px 16px;
            font-weight:600;
            color:#283247;
        }

        .info-body{
            padding:18px 22px;
            color:#2f3747;
            line-height:1.9;
            font-size:1rem;
        }

        div[data-testid="stLinkButton"] > a{
            width:100%;
            border-radius:10px;
            font-weight:600;
            border:1px solid #cfd5dd;
        }

    </style>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("## The Codex")

st.markdown('<div class="main-title">CODEX.ID</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="subtitle">Aplikasi maintainance all in one.</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="info-box">
        <div class="info-head">⌄ &nbsp; Kebutuhan File</div>
        <div class="info-body">
            <ul style="margin-top:0;padding-left:18px;">
                <li>Download dan install file <b>CodexSetup.exe</b>.</li>
                <li>Download file <b>AddOn</b> dalam format Excel.</li>
                <li>Setelah install selesai, shortcut aplikasi akan muncul di desktop.</li>
                <li>Jalankan aplikasi dan lanjutkan proses seperti biasa.</li>
            </ul>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("<div style='height:20px'></div>", unsafe_allow_html=True)

btn_col1, btn_col2, btn_col3 = st.columns([1,1,2])

with btn_col1:
    st.link_button("Download Codex", CODEX_URL, use_container_width=True)

with btn_col2:
    st.link_button("Download File AddOn", ADDON_URL, use_container_width=True)
