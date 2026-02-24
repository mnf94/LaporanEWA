import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import urllib.parse
from datetime import datetime
import pytz
import pandas as pd
import io

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="EWA Automator System", layout="wide", initial_sidebar_state="collapsed")

# --- KAMUS MAPPING BANK (XENDIT -> FINLINK) ---
BANK_MAPPING = {
    'ACEH': {'finlink_code': 'SYACIDJ1', 'finlink_name': 'BPD Aceh'}, 'AGRIS': {'finlink_code': 'AGTBIDJA', 'finlink_name': 'Bank Agris'},
    'AGRONIAGA': {'finlink_code': 'AGTBIDJA', 'finlink_name': 'Bank BRI Agroniaga'}, 'AMAR': {'finlink_code': 'LOMAIDJ1', 'finlink_name': 'Bank Amar Indonesia'},
    'ANZ': {'finlink_code': 'ANZBIDJX', 'finlink_name': 'Bank ANZ Indonesia'}, 'ARTHA': {'finlink_code': 'ARTGIDJA', 'finlink_name': 'Bank Artha Graha International'},
    'ARTOS': {'finlink_code': 'JAGBIDJA', 'finlink_name': 'Bank Artos Indonesia (Bank Jago)'}, 'BALI': {'finlink_code': 'ABALIDBS', 'finlink_name': 'BPD Bali'},
    'BAML': {'finlink_code': 'BOFAID2X', 'finlink_name': 'Bank of America Merill-Lynch'}, 'BANTEN': {'finlink_code': 'PDBBIDJ1', 'finlink_name': 'BPD Banten'},
    'BCA': {'finlink_code': 'CENAIDJA', 'finlink_name': 'Bank Central Asia (BCA)'}, 'BCA_DIGITAL': {'finlink_code': 'BBLUIDJA', 'finlink_name': 'BCA Digital (Blu)'},
    'BCA_SYR': {'finlink_code': 'BSYAIDJA', 'finlink_name': 'Bank Central Asia (BCA) Syariah'}, 'BENGKULU': {'finlink_code': 'PDBKIDJ1', 'finlink_name': 'BPD Bengkulu'},
    'BJB': {'finlink_code': 'PDJBIDJA', 'finlink_name': 'Bank BJB'}, 'BJB_SYR': {'finlink_code': 'SYJBIDJ1', 'finlink_name': 'Bank BJB Syariah'},
    'BNI': {'finlink_code': 'BNINIDJA', 'finlink_name': 'Bank Negara Indonesia (BNI)'}, 'BNP_PARIBAS': {'finlink_code': 'BNPAIDJA', 'finlink_name': 'Bank BNP Paribas'},
    'BOC': {'finlink_code': 'BKCHIDJA', 'finlink_name': 'Bank of China (BOC)'}, 'BPD_KALTIMTARA': {'finlink_code': 'PDKTIDJ1', 'finlink_name': 'BPD Kaltim Kaltara'},
    'BRI': {'finlink_code': 'BRINIDJA', 'finlink_name': 'Bank Rakyat Indonesia (BRI)'}, 'BSI': {'finlink_code': 'BSMDIDJA', 'finlink_name': 'Bank Syariah Indonesia'},
    'BTN': {'finlink_code': 'BTANIDJA', 'finlink_name': 'Bank Tabungan Negara (BTN)'}, 'BTN_UUS': {'finlink_code': 'SYBTIDJ1', 'finlink_name': 'Bank Tabungan Negara (BTN) UUS'},
    'BTPN_SYARIAH': {'finlink_code': 'PUBAIDJ1', 'finlink_name': 'BTPN Syariah'}, 'BUKOPIN': {'finlink_code': 'BBUKIDJA', 'finlink_name': 'Bank Bukopin'},
    'BUKOPIN_SYR': {'finlink_code': 'SDOBIDJ1', 'finlink_name': 'Bank Syariah Bukopin'}, 'BUMI_ARTA': {'finlink_code': 'BBAIDJA', 'finlink_name': 'Bank Bumi Arta'},
    'CAPITAL': {'finlink_code': 'BCIAIDJA', 'finlink_name': 'Bank Capital Indonesia'}, 'CCB': {'finlink_code': 'MCORIDJA', 'finlink_name': 'China Construction Bank Indonesia'},
    'CHINATRUST': {'finlink_code': 'CTCBIDJA', 'finlink_name': 'Bank Chinatrust Indonesia'}, 'CIMB': {'finlink_code': 'BNIAIDJA', 'finlink_name': 'Bank CIMB Niaga'},
    'CIMB_UUS': {'finlink_code': 'SYNAIDJ1', 'finlink_name': 'Bank CIMB Niaga UUS'}, 'CITIBANK': {'finlink_code': 'CITIIDJX', 'finlink_name': 'Citibank'},
    'DAERAH_ISTIMEWA': {'finlink_code': 'PDYKIDJ1', 'finlink_name': 'BPD Daerah Istimewa Yogyakarta (DIY)'}, 'DAERAH_ISTIMEWA_UUS': {'finlink_code': 'SYYKIDJ1', 'finlink_name': 'BPD Daerah Istimewa Yogyakarta (DIY) UUS'},
    'DANA': {'finlink_code': 'DANAIDJ1', 'finlink_name': 'DANA eWallet'}, 'DANAMON': {'finlink_code': 'BDINIDJA', 'finlink_name': 'Bank Danamon'},
    'DANAMON_UUS': {'finlink_code': 'SYBDIDJ1', 'finlink_name': 'Bank Danamon UUS'}, 'DBS': {'finlink_code': 'DBSBIDJA', 'finlink_name': 'Bank DBS Indonesia'},
    'DEUTSCHE': {'finlink_code': 'DEUTIDJA', 'finlink_name': 'Deutsche Bank'}, 'DKI': {'finlink_code': 'BDKIIDJ1', 'finlink_name': 'Bank DKI'},
    'DKI_UUS': {'finlink_code': 'SYDKIDJ1', 'finlink_name': 'Bank DKI UUS'}, 'FAMA': {'finlink_code': 'FAMAIDJ1', 'finlink_name': 'Bank Fama International'},
    'GANESHA': {'finlink_code': 'GNESIDJA', 'finlink_name': 'Bank Ganesha'}, 'HANA': {'finlink_code': 'HNBNIDJA', 'finlink_name': 'Bank Hana'},
    'HSBC': {'finlink_code': 'HSBCIDJA', 'finlink_name': 'Hongkong and Shanghai Bank (HSBC)'}, 'ICBC': {'finlink_code': 'ICBKIDJA', 'finlink_name': 'Bank ICBC Indonesia'},
    'INA_PERDANA': {'finlink_code': 'IAPTIDJA', 'finlink_name': 'Bank Ina Perdania'}, 'INDEX_SELINDO': {'finlink_code': 'BIDXIDJA', 'finlink_name': 'Bank Index Selindo'},
    'INDIA': {'finlink_code': 'BKIDIDJA', 'finlink_name': 'Bank of India Indonesia'}, 'JAMBI': {'finlink_code': 'PDJMIDJ1', 'finlink_name': 'BPD Jambi'},
    'JAMBI_UUS': {'finlink_code': 'SYJMIDJ1', 'finlink_name': 'BPD Jambi UUS'}, 'JASA_JAKARTA': {'finlink_code': 'JSABIDJ1', 'finlink_name': 'Bank Jasa Jakarta'},
    'JAWA_TENGAH': {'finlink_code': 'PDJGIDJ1', 'finlink_name': 'BPD Jawa Tengah'}, 'JAWA_TENGAH_UUS': {'finlink_code': 'SYJGIDJ1', 'finlink_name': 'BPD Jawa Tengah UUS'},
    'JAWA_TIMUR': {'finlink_code': 'PDJTIDJ1', 'finlink_name': 'BPD Jawa Timur'}, 'JAWA_TIMUR_UUS': {'finlink_code': 'SYJTIDJ1', 'finlink_name': 'BPD Jawa Timur UUS'},
    'JPMORGAN': {'finlink_code': 'CHASIDJX', 'finlink_name': 'JP Morgan Chase Bank'}, 'JTRUST': {'finlink_code': 'CICTIDJA', 'finlink_name': 'Bank JTrust Indonesia'},
    'KALIMANTAN_BARAT': {'finlink_code': 'PDKBIDJ1', 'finlink_name': 'BPD Kalimantan Barat'}, 'KALIMANTAN_BARAT_UUS': {'finlink_code': 'SYKBIDJ1', 'finlink_name': 'BPD Kalimantan Barat UUS'},
    'KALIMANTAN_SELATAN': {'finlink_code': 'PDKSIDJ1', 'finlink_name': 'BPD Kalimantan Selatan'}, 'KALIMANTAN_SELATAN_UUS': {'finlink_code': 'SYKSIDJ1', 'finlink_name': 'BPD Kalimantan Selatan UUS'},
    'KALIMANTAN_TENGAH': {'finlink_code': 'PDGKIDJ1', 'finlink_name': 'BPD Kalimantan Tengah'}, 'KALIMANTAN_TIMUR': {'finlink_code': 'PDKTIDJ1', 'finlink_name': 'BPD Kalimantan Timur'},
    'KALIMANTAN_TIMUR_UUS': {'finlink_code': 'SYKTIDJ1', 'finlink_name': 'BPD Kalimantan Timur UUS'}, 'KESEJAHTERAAN_EKONOMI': {'finlink_code': 'SSPIIDJA', 'finlink_name': 'Bank Kesejahteraan Ekonomi'},
    'LAMPUNG': {'finlink_code': 'PDLPIDJ1', 'finlink_name': 'BPD Lampung'}, 'MALUKU': {'finlink_code': 'PDMLIDJ1', 'finlink_name': 'BPD Maluku'},
    'MANDIRI': {'finlink_code': 'BMRIIDJA', 'finlink_name': 'Bank Mandiri'}, 'MANDIRI_TASPEN': {'finlink_code': 'SIHIDJ1', 'finlink_name': 'Mandiri Taspen Pos'},
    'MAYAPADA': {'finlink_code': 'MAYAIDJA', 'finlink_name': 'Bank Mayapada International'}, 'MAYBANK': {'finlink_code': 'IBBKIDJA', 'finlink_name': 'Bank Maybank'},
    'MAYBANK_SYR': {'finlink_code': 'SYBKIDJ1', 'finlink_name': 'Bank Maybank Syariah Indonesia'}, 'MEGA': {'finlink_code': 'MEGAIDJA', 'finlink_name': 'Bank Mega'},
    'MEGA_SYR': {'finlink_code': 'BUTGIDJ1', 'finlink_name': 'Bank Syariah Mega'}, 'MESTIKA_DHARMA': {'finlink_code': 'MEDHIDS1', 'finlink_name': 'Bank Mestika Dharma'},
    'MITSUI': {'finlink_code': 'SUNIIDJA', 'finlink_name': 'Bank Sumitomo Mitsui Indonesia'}, 'MIZUHO': {'finlink_code': 'MHCCIDJA', 'finlink_name': 'Bank Mizuho Indonesia'},
    'MNC_INTERNASIONAL': {'finlink_code': 'BUMIIDJA', 'finlink_name': 'Bank MNC Internasional'}, 'MUAMALAT': {'finlink_code': 'MUABIDJA', 'finlink_name': 'Bank Muamalat Indonesia'},
    'MULTI_ARTA_SENTOSA': {'finlink_code': 'BMSEIDJA', 'finlink_name': 'Bank Multi Arta Sentosa'}, 'NATIONALNOBU': {'finlink_code': 'LFIBIDJ1', 'finlink_name': 'Bank Nationalnobu'},
    'NUSA_TENGGARA_TIMUR': {'finlink_code': 'PDNTIDJA', 'finlink_name': 'BPD Nusa Tenggara Timur'}, 'OKE': {'finlink_code': 'LMANIDJ1', 'finlink_name': 'Bank Oke Indonesia'},
    'PANIN': {'finlink_code': 'PINBIDJA', 'finlink_name': 'Bank Panin'}, 'PANIN_SYR': {'finlink_code': 'ARFAIDJ1', 'finlink_name': 'Bank Panin Syariah'},
    'PAPUA': {'finlink_code': 'PDIJIDJ1', 'finlink_name': 'BPD Papua'}, 'PERMATA': {'finlink_code': 'BBBAIDJA', 'finlink_name': 'Bank Permata'},
    'PERMATA_UUS': {'finlink_code': 'SYBBIDJ1', 'finlink_name': 'Bank Permata UUS'}, 'QNB_INDONESIA': {'finlink_code': 'AWANIDJA', 'finlink_name': 'Bank QNB Indonesia'},
    'RESONA': {'finlink_code': 'BPIAIDJA', 'finlink_name': 'Bank Resona Perdania'}, 'RIAU_DAN_KEPRI': {'finlink_code': 'PDRIIDJA', 'finlink_name': 'BPD Riau Dan Kepri'},
    'SAHABAT_SAMPOERNA': {'finlink_code': 'SAHMIDJA', 'finlink_name': 'Bank Sahabat Sampoerna'}, 'SBI_INDONESIA': {'finlink_code': 'SBIDIDJA', 'finlink_name': 'Bank SBI Indonesia'},
    'SHINHAN': {'finlink_code': 'MEEKIDJ1', 'finlink_name': 'Bank Shinhan Indonesia'}, 'SHOPEEPAY': {'finlink_code': 'APIDIDJ1', 'finlink_name': 'Shopee Pay eWallet'},
    'SINARMAS': {'finlink_code': 'SBJKIDJA', 'finlink_name': 'Sinarmas'}, 'STANDARD_CHARTERED': {'finlink_code': 'SCBLIDJX', 'finlink_name': 'Standard Charted Bank'},
    'SULAWESI': {'finlink_code': 'PDWGIDJ1', 'finlink_name': 'BPD Sulawesi Tengah'}, 'SULAWESI_TENGGARA': {'finlink_code': 'PDWRIDJ1', 'finlink_name': 'BPD Sulawesi Tenggara'},
    'SULSELBAR': {'finlink_code': 'PDWSIDJA', 'finlink_name': 'BPD Sulselbar'}, 'SULSELBAR_UUS': {'finlink_code': 'SYWSIDJ1', 'finlink_name': 'BPD Sulselbar UUS'},
    'SULUT': {'finlink_code': 'PDWUIDJ1', 'finlink_name': 'BPD Sulut'}, 'SUMATERA_BARAT': {'finlink_code': 'PDSBIDJ1', 'finlink_name': 'BPD Sumatera Barat'},
    'SUMATERA_BARAT_UUS': {'finlink_code': 'SYSBIDJ1', 'finlink_name': 'BPD Sumatera Barat UUS'}, 'SUMSEL_DAN_BABEL': {'finlink_code': 'BSSPIDSP', 'finlink_name': 'BPD Sumsel Dan Babel'},
    'SUMSEL_DAN_BABEL_UUS': {'finlink_code': 'SYSSIDJ1', 'finlink_name': 'BPD Sumsel Dan Babel UUS'}, 'SUMUT': {'finlink_code': 'PDSUIDJ1', 'finlink_name': 'BPD Sumut'},
    'SUMUT_UUS': {'finlink_code': 'SYSUIDJ1', 'finlink_name': 'BPD Sumut UUS'}, 'TOKYO': {'finlink_code': 'BOTKIDJX', 'finlink_name': 'Bank of Tokyo Mitsubishi UFJ'},
    'UOB': {'finlink_code': 'BBJIDJA', 'finlink_name': 'Bank UOB Indonesia'}, 'VICTORIA_INTERNASIONAL': {'finlink_code': 'VICTIDJ1', 'finlink_name': 'Bank Victoria Internasional'},
    'VICTORIA_SYR': {'finlink_code': 'SWAGIDJ1', 'finlink_name': 'Bank Victoria Syariah'}, 'WOORI': {'finlink_code': 'BSDRIDJA', 'finlink_name': 'Bank Woori Saudara Indonesia'},
    'YUDHA_BHAKTI': {'finlink_code': 'YUDBIDJ1', 'finlink_name': 'Bank Yudha Bhakti'}
}

# --- CSS: AURORA, GLASSMORPHISM, SAAS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Outfit', sans-serif !important; }
    .stApp { background: linear-gradient(-45deg, #e0f2fe, #d1fae5, #ffffff, #bae6fd); background-size: 400% 400%; animation: gradientBG 15s ease infinite; }
    @keyframes gradientBG { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
    [data-testid="stForm"], .glass-container {
        background: rgba(255, 255, 255, 0.35) !important; backdrop-filter: blur(16px) !important; -webkit-backdrop-filter: blur(16px) !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important; border-radius: 24px !important; padding: 30px !important; box-shadow: 0 10px 40px rgba(0, 0, 0, 0.08) !important;
    }
    div[data-baseweb="input"] > div { background-color: rgba(255, 255, 255, 0.6) !important; border-radius: 12px !important; border: 1px solid rgba(255, 255, 255, 0.8) !important; transition: all 0.3s ease; }
    div[data-baseweb="input"] > div:focus-within { background-color: #ffffff !important; border-color: #0ea5e9 !important; box-shadow: 0 0 0 3px rgba(14, 165, 233, 0.2) !important; }
    [data-testid="baseButton-primary"] { background: linear-gradient(135deg, #0ea5e9 0%, #f97316 100%) !important; color: white !important; border: none !important; border-radius: 12px !important; font-size: 16px !important; font-weight: 600 !important; padding: 10px 24px !important; box-shadow: 0 4px 15px rgba(14, 165, 233, 0.3) !important; transition: transform 0.2s ease, box-shadow 0.2s ease !important; }
    [data-testid="baseButton-primary"]:hover { transform: translateY(-2px) !important; box-shadow: 0 6px 20px rgba(14, 165, 233, 0.4) !important; }
</style>
""", unsafe_allow_html=True)

# --- FUNGSI HELPER WAKTU & ANGKA ---
def get_current_time_wib():
    return datetime.now(pytz.timezone('Asia/Jakarta')).strftime("%d %B %Y - %H:%M:%S WIB")

def format_rp(angka):
    try:
        if angka is None or angka == "": return "Rp 0"
        return f"Rp {int(round(float(angka))):,}".replace(",", ".")
    except: return "Rp 0"

def format_num(angka):
    try:
        if angka is None or angka == "": return "0"
        return f"{int(round(float(angka))):,}".replace(",", ".")
    except: return "0"

def safe_int(value, default=0):
    try:
        if value is None or str(value).strip() == "": return default
        return int(float(value))
    except: return default

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Form' if 'bank_code' in df.columns else 'Template')
    return output.getvalue()

# --- FUNGSI KONEKSI GOOGLE SHEETS ---
def connect_google_sheets():
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        return gspread.authorize(creds).open("Database_EWA_App").sheet1 
    except: return None

# --- LOGIKA LOGIN SISTEM ---
def check_login():
    if "logged_in" not in st.session_state: st.session_state["logged_in"] = False
    if not st.session_state["logged_in"]:
        _, col2, _ = st.columns([1, 2, 1])
        with col2:
            st.markdown("<br><br><br><h2 style='text-align: center; color: #0ea5e9;'>🔐 Login EWA Automator</h2>", unsafe_allow_html=True)
            with st.form("login_form"):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                if st.form_submit_button("Masuk", type="primary"):
                    if "users" in st.secrets and username in st.secrets["users"] and st.secrets["users"][username] == password:
                        st.session_state["logged_in"] = True; st.session_state["username"] = username; st.rerun()
                    else: st.error("❌ Username/Password salah!")
        return False
    return True

# --- MENJALANKAN APLIKASI UTAMA ---
if check_login():
    
    # NAVBAR
    col_title, col_menu, col_logout = st.columns([5, 2, 1])
    with col_title:
        st.title("✨ EWA Automator System")
    with col_menu:
        st.markdown("<br>", unsafe_allow_html=True)
        menu_selection = st.selectbox("Menu", ["📊 Laporan Dashboard", "⚙️ Proses EWA (Mapping)"], label_visibility="collapsed")
    with col_logout:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True):
            st.session_state["logged_in"] = False; st.session_state["username"] = ""; st.rerun()

    # --- MENU 1: LAPORAN DASHBOARD ---
    if menu_selection == "📊 Laporan Dashboard":
        db_sheet = connect_google_sheets()
        history = {}
        if db_sheet is None: st.warning("⚠️ Google Sheets Belum Terhubung. Form berjalan tanpa histori.")
        else:
            st.success("✅ Terhubung ke Google Sheets! Data histori otomatis ditarik.")
            try:
                records = db_sheet.get_all_records()
                history = records[-1] if records else {}
            except: pass
            if history.get("created_by", "-") != "-": st.info(f"🕒 **Laporan sebelumnya digenerate oleh:** `{history.get('created_by')}` | **Waktu:** `{history.get('timestamp')}`")

        with st.form("report_form"):
            st.markdown("### 📅 Informasi Umum")
            c1, c2 = st.columns(2)
            with c1: tanggal_laporan = st.text_input("Tanggal Laporan (Muncul di HTML)", value=datetime.now().strftime('%d %B %Y'))
            with c2: hari_berjalan = st.number_input("Hari Berjalan Bulan Ini (Utk Proyeksi)", min_value=1, value=int(datetime.now().strftime('%d')), format="%d")

            st.markdown("### 📊 1. Report Hari Ini")
            col1, col2, col3 = st.columns(3)
            with col1:
                qty_ewa = st.number_input("Total Qty EWA", value=None, min_value=0, format="%d", step=1)
                total_ewa = st.number_input("Total EWA (Rp)", value=None, min_value=0, format="%d", step=1)
                admin_ewa = st.number_input("Total Biaya Admin (Rp)", value=None, min_value=0, format="%d", step=1)
                transfer_ewa = st.number_input("Total Biaya Transfer (Rp)", value=None, min_value=0, format="%d", step=1)
            with col2:
                profit_ewa = st.number_input("Total Profit (Rp)", value=None, min_value=0, format="%d", step=1)
                xendit_ewa = st.number_input("Total Transfer Xendit (Rp)", value=None, min_value=0, format="%d", step=1)
                finlink_ewa = st.number_input("Total Transfer Finlink (Rp)", value=None, min_value=0, format="%d", step=1)
                pending_ewa = st.number_input("EWA Pending (Qty)", value=None, min_value=0, format="%d", step=1)
            with col3:
                qty_ppob = st.number_input("Total Qty EWA PPOB", value=None, min_value=0, format="%d", step=1)
                ewa_ppob = st.number_input("Total EWA PPOB (Rp)", value=None, min_value=0, format="%d", step=1)
                admin_ppob = st.number_input("Total Biaya Admin PPOB (Rp)", value=None, min_value=0, format="%d", step=1)

            st.markdown("### 💳 2. Saldo PPOB & Payment Gateway")
            col4, col5 = st.columns(2)
            with col4:
                saldo_alterra = st.number_input("Saldo PPOB Alterra (Rp)", value=None, min_value=0, format="%d", step=1)
                saldo_pelangi = st.number_input("Saldo PPOB Pelangi (Rp)", value=None, min_value=0, format="%d", step=1)
                saldo_uv = st.number_input("Saldo PPOB Ultra Voucher (Rp)", value=None, min_value=0, format="%d", step=1)
            with col5:
                saldo_xendit = st.number_input("Saldo PG Xendit (Rp)", value=None, min_value=0, format="%d", step=1)
                saldo_finlink = st.number_input("Saldo PG Finlink (Rp)", value=None, min_value=0, format="%d", step=1)

            st.markdown("### 📈 3. Report Bulan Berjalan (MTD)")
            mtd_periode = st.text_input("Periode MTD", value=history.get("mtd_periode", "1 - 20 Februari 2026"))
            col6, col7, col8 = st.columns(3)
            with col6:
                mtd_qty_ewa = st.number_input("MTD Qty EWA", value=safe_int(history.get("mtd_qty_ewa", 0)), min_value=0, format="%d")
                mtd_total_ewa = st.number_input("MTD Total EWA (Rp)", value=safe_int(history.get("mtd_total_ewa", 0)), min_value=0, format="%d")
                mtd_admin = st.number_input("MTD Biaya Admin (Rp)", value=safe_int(history.get("mtd_admin", 0)), min_value=0, format="%d")
                mtd_transfer = st.number_input("MTD Biaya Transfer (Rp)", value=safe_int(history.get("mtd_transfer", 0)), min_value=0, format="%d")
            with col7:
                mtd_profit = st.number_input("MTD Profit (Rp)", value=safe_int(history.get("mtd_profit", 0)), min_value=0, format="%d")
                mtd_xendit = st.number_input("MTD Trf Xendit (Rp)", value=safe_int(history.get("mtd_xendit", 0)), min_value=0, format="%d")
                mtd_finlink = st.number_input("MTD Trf Finlink (Rp)", value=safe_int(history.get("mtd_finlink", 0)), min_value=0, format="%d")
            with col8:
                mtd_qty_ppob = st.number_input("MTD Qty PPOB", value=safe_int(history.get("mtd_qty_ppob", 0)), min_value=0, format="%d")
                mtd_ewa_ppob = st.number_input("MTD EWA PPOB (Rp)", value=safe_int(history.get("mtd_ewa_ppob", 0)), min_value=0, format="%d")
                mtd_admin_ppob = st.number_input("MTD Admin PPOB (Rp)", value=safe_int(history.get("mtd_admin_ppob", 0)), min_value=0, format="%d")

            st.markdown("### 🔙 4. Report Bulan Lalu")
            lm_periode = st.text_input("Periode Bulan Lalu", value=history.get("lm_periode", "1 - 31 Januari 2026"))
            col9, col10, col11 = st.columns(3)
            with col9:
                lm_qty_ewa = st.number_input("LM Qty EWA", value=safe_int(history.get("lm_qty_ewa", 0)), min_value=0, format="%d")
                lm_total_ewa = st.number_input("LM Total EWA (Rp)", value=safe_int(history.get("lm_total_ewa", 0)), min_value=0, format="%d")
                lm_admin = st.number_input("LM Biaya Admin (Rp)", value=safe_int(history.get("lm_admin", 0)), min_value=0, format="%d")
                lm_transfer = st.number_input("LM Biaya Transfer (Rp)", value=safe_int(history.get("lm_transfer", 0)), min_value=0, format="%d")
            with col10:
                lm_profit = st.number_input("LM Profit (Rp)", value=safe_int(history.get("lm_profit", 0)), min_value=0, format="%d")
                lm_xendit = st.number_input("LM Trf Xendit (Rp)", value=safe_int(history.get("lm_xendit", 0)), min_value=0, format="%d")
                lm_finlink = st.number_input("LM Trf Finlink (Rp)", value=safe_int(history.get("lm_finlink", 0)), min_value=0, format="%d")
            with col11:
                lm_qty_ppob = st.number_input("LM Qty PPOB", value=safe_int(history.get("lm_qty_ppob", 0)), min_value=0, format="%d")
                lm_ewa_ppob = st.number_input("LM EWA PPOB (Rp)", value=safe_int(history.get("lm_ewa_ppob", 0)), min_value=0, format="%d")
                lm_admin_ppob = st.number_input("LM Admin PPOB (Rp)", value=safe_int(history.get("lm_admin_ppob", 0)), min_value=0, format="%d")

            submitted = st.form_submit_button("🚀 Buat Laporan HTML Sekarang!", type="primary")

        if submitted:
            qty_ewa = safe_int(qty_ewa); total_ewa = safe_int(total_ewa); admin_ewa = safe_int(admin_ewa); transfer_ewa = safe_int(transfer_ewa)
            profit_ewa = safe_int(profit_ewa); xendit_ewa = safe_int(xendit_ewa); finlink_ewa = safe_int(finlink_ewa); pending_ewa = safe_int(pending_ewa)
            qty_ppob = safe_int(qty_ppob); ewa_ppob = safe_int(ewa_ppob); admin_ppob = safe_int(admin_ppob); saldo_alterra = safe_int(saldo_alterra)
            saldo_pelangi = safe_int(saldo_pelangi); saldo_uv = safe_int(saldo_uv); saldo_xendit = safe_int(saldo_xendit); saldo_finlink = safe_int(saldo_finlink)

            if db_sheet is not None:
                new_hist = {"created_by": st.session_state["username"], "timestamp": get_current_time_wib(), "mtd_periode": mtd_periode, "mtd_qty_ewa": mtd_qty_ewa, "mtd_total_ewa": mtd_total_ewa, "mtd_admin": mtd_admin, "mtd_transfer": mtd_transfer, "mtd_profit": mtd_profit, "mtd_xendit": mtd_xendit, "mtd_finlink": mtd_finlink, "mtd_qty_ppob": mtd_qty_ppob, "mtd_ewa_ppob": mtd_ewa_ppob, "mtd_admin_ppob": mtd_admin_ppob, "lm_periode": lm_periode, "lm_qty_ewa": lm_qty_ewa, "lm_total_ewa": lm_total_ewa, "lm_admin": lm_admin, "lm_transfer": lm_transfer, "lm_profit": lm_profit, "lm_xendit": lm_xendit, "lm_finlink": lm_finlink, "lm_qty_ppob": lm_qty_ppob, "lm_ewa_ppob": lm_ewa_ppob, "lm_admin_ppob": lm_admin_ppob}
                try:
                    if len(db_sheet.get_all_values()) == 0: db_sheet.append_row(list(new_hist.keys()))
                    db_sheet.append_row(list(new_hist.values()))
                except: pass

            avg_ewa_per_trx = total_ewa / qty_ewa if qty_ewa > 0 else 0
            fee_per_trx = profit_ewa / qty_ewa if qty_ewa > 0 else 0
            margin_pct = (profit_ewa / total_ewa * 100) if total_ewa > 0 else 0
            pending_pct = (pending_ewa / qty_ewa * 100) if qty_ewa > 0 else 0
            
            # --- INI BARIS YANG HILANG DAN SUDAH DIKEMBALIKAN ---
            margin_mtd_pct = (mtd_profit / mtd_total_ewa * 100) if mtd_total_ewa > 0 else 0
            
            avg_daily_mtd_total = mtd_total_ewa / hari_berjalan if hari_berjalan > 0 else 0
            avg_daily_lm_total = lm_total_ewa / 31 
            avg_daily_mtd_profit = mtd_profit / hari_berjalan if hari_berjalan > 0 else 0
            diff_total_ewa_mtd_lm = ((avg_daily_mtd_total - avg_daily_lm_total) / avg_daily_lm_total * 100) if avg_daily_lm_total > 0 else 0
            diff_profit_daily_mtd = ((profit_ewa - avg_daily_mtd_profit) / avg_daily_mtd_profit * 100) if avg_daily_mtd_profit > 0 else 0
            runrate_total_ewa = avg_daily_mtd_total * 30; runrate_profit = avg_daily_mtd_profit * 30

            growth_arrow = "🟢 <b>Naik</b>" if diff_total_ewa_mtd_lm >= 0 else "🔴 <b>Turun</b>"
            exec_summary = f"Performa rata-rata harian MTD sebesar {format_rp(avg_daily_mtd_total)}, {growth_arrow} <b>{abs(diff_total_ewa_mtd_lm):.1f}%</b> dibandingkan rata-rata harian bulan lalu. Proyeksi EWA bulan ini mencapai {format_rp(runrate_total_ewa)}."

            margin_badge = '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#dcfce7;color:#166534;font-size:11px;font-weight:700;">Hijau</span>' if margin_pct >= 4 else '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;">Kuning</span>' if margin_pct >= 3 else '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#fee2e2;color:#991b1b;font-size:11px;font-weight:700;">Merah</span>'
            pending_badge = '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#dcfce7;color:#166534;font-size:11px;font-weight:700;">Aman</span>' if pending_pct < 10 else '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#fef3c7;color:#92400e;font-size:11px;font-weight:700;">Waspada</span>' if pending_pct < 20 else '<span style="display:inline-block;margin-left:8px;padding:2px 8px;border-radius:999px;background:#fee2e2;color:#991b1b;font-size:11px;font-weight:700;">Bahaya</span>'
            alert_box = f'<tr><td style="padding:0 24px 16px 24px;"><table width="100%" style="background:#fee2e2;border:1px solid #ef4444;border-radius:12px;"><tr><td style="padding:12px 14px;color:#991b1b;font-weight:700;font-size:14px;">🚨 PERINGATAN KRITIS: Rasio Pending Tinggi</td></tr><tr><td style="padding:0 14px 12px 14px;color:#7f1d1d;font-size:13px;">Terdapat <b>{format_num(pending_ewa)} transaksi pending</b> dari total {format_num(qty_ewa)} transaksi EWA hari ini. Rasio menembus <b>{pending_pct:.2f}%</b>. Mohon periksa status gateway.</td></tr></table></td></tr>' if pending_pct >= 20 else ""

            js_func_rp = r"function(v) { if (!v) return 'Rp 0'; return 'Rp ' + v.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }"
            js_func_qty = r"function(v) { if (!v) return '0'; return v.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }"
            js_func_pie = r"function(v, ctx) { if (!v) return ''; let sum = 0; let dataArr = ctx.chart.data.datasets[0].data; dataArr.map(data => { sum += data; }); let pct = (v*100 / sum).toFixed(1) + '%'; let val = v.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); return pct + '\nRp ' + val; }"

            c1_json = json.dumps({"type": "bar", "data": {"labels": ["Total EWA", "Profit", "Trf Xendit", "EWA PPOB"], "datasets": [{"label": "MTD", "data": [mtd_total_ewa, mtd_profit, mtd_xendit, mtd_ewa_ppob], "backgroundColor": "#0ea5e9", "borderRadius": 4}, {"label": "Bulan Lalu", "data": [lm_total_ewa, lm_profit, lm_xendit, lm_ewa_ppob], "backgroundColor": "#f97316", "borderRadius": 4}]}, "options": {"layout": {"padding": {"top": 35}}, "plugins": {"datalabels": { "display": True, "align": "end", "anchor": "end", "color": "#0f172a", "font": {"size": 10, "weight": "bold"}, "formatter": "FORMATTER_RP" }}, "legend": {"position": "bottom", "labels": {"usePointStyle": True, "padding": 20}}, "scales": { "yAxes": [{"display": False}], "xAxes": [{"gridLines": {"display": False}, "ticks": {"fontStyle": "bold"}}] }}}).replace('"FORMATTER_RP"', js_func_rp)
            chart1_url = f"https://quickchart.io/chart?w=600&h=450&devicePixelRatio=2&c={urllib.parse.quote(c1_json)}"

            c3_json = json.dumps({"type": "doughnut", "data": {"labels": ["Profit EWA", "Profit PPOB"], "datasets": [{"data": [mtd_profit, mtd_admin_ppob], "backgroundColor": ["#0ea5e9", "#f97316"]}]}, "options": {"plugins": {"datalabels": { "display": True, "color": "#fff", "font": {"weight": "bold", "size": 12}, "formatter": "FORMATTER_PIE", "textAlign": "center" }}, "legend": {"position": "bottom", "labels": {"usePointStyle": True, "padding": 20}}}}).replace('"FORMATTER_PIE"', js_func_pie)
            chart3_url = f"https://quickchart.io/chart?w=600&h=450&devicePixelRatio=2&c={urllib.parse.quote(c3_json)}"

            c2_json = json.dumps({"type": "bar", "data": {"labels": ["Qty EWA", "Qty PPOB"], "datasets": [{"label": "MTD", "data": [mtd_qty_ewa, mtd_qty_ppob], "backgroundColor": "#0ea5e9", "borderRadius": 4}, {"label": "Bulan Lalu", "data": [lm_qty_ewa, lm_qty_ppob], "backgroundColor": "#f97316", "borderRadius": 4}]}, "options": {"layout": {"padding": {"top": 35}}, "plugins": {"datalabels": { "display": True, "align": "end", "anchor": "end", "color": "#0f172a", "font": {"size": 12, "weight": "bold"}, "formatter": "FORMATTER_QTY" }}, "legend": {"position": "bottom", "labels": {"usePointStyle": True, "padding": 20}}, "scales": { "yAxes": [{"display": False}], "xAxes": [{"gridLines": {"display": False}, "ticks": {"fontStyle": "bold"}}] }}}).replace('"FORMATTER_QTY"', js_func_qty)
            chart2_url = f"https://quickchart.io/chart?w=800&h=400&devicePixelRatio=2&c={urllib.parse.quote(c2_json)}"

            html_output = f"""<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8"/><meta name="viewport" content="width=device-width, initial-scale=1.0"/><title>Laporan EWA & PPOB</title></head><body style="margin:0;padding:0;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;color:#0f172a;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#f1f5f9;margin:0;padding:24px 0;"><tr><td align="center"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:980px;background:#f1f5f9;"><tr><td style="padding:0 24px 16px 24px;"><table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:linear-gradient(135deg,#0ea5e9 0%,#f97316 100%);border-radius:14px;box-shadow:0 4px 12px rgba(14,165,233,0.2);"><tr><td style="padding:20px 22px;color:#ffffff;"><div style="font-size:22px;font-weight:700;line-height:1.2;">Laporan EWA &amp; PPOB</div><div style="font-size:13px;opacity:.95;margin-top:6px;">Tanggal: {tanggal_laporan}</div></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;border-left:4px solid #0ea5e9;"><tr><td style="padding:14px 16px;color:#0f172a;font-size:14px;line-height:1.6;">💡 <b>Executive Summary:</b><br>{exec_summary}</td></tr></table></td></tr><tr><td style="padding:0 24px 10px 24px;"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="padding:0 6px 10px 0;width:25%;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;"><tr><td style="padding:12px 12px 4px;color:#64748b;font-size:12px;">Qty</td></tr><tr><td style="padding:0 12px 12px;color:#0f172a;font-size:20px;font-weight:700;">{format_num(qty_ewa)}</td></tr></table></td><td style="padding:0 6px 10px 6px;width:25%;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;"><tr><td style="padding:12px 12px 4px;color:#64748b;font-size:12px;">Total EWA</td></tr><tr><td style="padding:0 12px 12px;color:#0f172a;font-size:20px;font-weight:700;">{format_rp(total_ewa)}</td></tr></table></td><td style="padding:0 6px 10px 6px;width:25%;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;"><tr><td style="padding:12px 12px 4px;color:#64748b;font-size:12px;">Profit</td></tr><tr><td style="padding:0 12px 12px;color:#0f172a;font-size:20px;font-weight:700;">{format_rp(profit_ewa)}</td></tr></table></td><td style="padding:0 0 10px 6px;width:25%;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;"><tr><td style="padding:12px 12px 4px;color:#64748b;font-size:12px;">Pending</td></tr><tr><td style="padding:0 12px 12px;color:#0f172a;font-size:20px;font-weight:700;">{format_num(pending_ewa)}</td></tr></table></td></tr></table></td></tr>{alert_box}<tr><td style="padding:0 24px 16px 24px;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">KPI &amp; Rasio Tambahan</td></tr><tr><td style="padding:12px;"><table width="100%" style="border-collapse:collapse;"><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Avg EWA per Transaksi</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(avg_ewa_per_trx)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Fee/Profit per Transaksi</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(fee_per_trx)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Margin Profit relatif</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{margin_pct:.2f}% {margin_badge}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Rasio Pending</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{pending_pct:.2f}% {pending_badge}</td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Report Harian</td></tr><tr><td style="padding:12px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(qty_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Total EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(total_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Admin</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(admin_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Transfer</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(transfer_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Profit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(profit_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Xendit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(xendit_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Finlink</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(finlink_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">EWA Pending</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(pending_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(qty_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">EWA PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(ewa_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Admin PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(admin_ppob)}</td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Saldo PPOB & Payment Gateway</td></tr><tr><td style="padding:12px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Saldo Alterra</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(saldo_alterra)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Saldo Pelangi</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(saldo_pelangi)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Saldo Ultra Voucher</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(saldo_uv)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Saldo Payment Gateway Xendit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(saldo_xendit)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Saldo Payment Gateway Finlink</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(saldo_finlink)}</td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Report Bulan Berjalan (MTD)</td></tr><tr><td style="padding:12px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Periode</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{mtd_periode}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(mtd_qty_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Total EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_total_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Admin</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_admin)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Transfer</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_transfer)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Profit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_profit)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Xendit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_xendit)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Finlink</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_finlink)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(mtd_qty_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">EWA PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_ewa_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Admin PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(mtd_admin_ppob)}</td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Report Bulan Lalu</td></tr><tr><td style="padding:12px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Periode</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{lm_periode}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(lm_qty_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Total EWA</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_total_ewa)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Admin</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_admin)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Biaya Transfer</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_transfer)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Profit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_profit)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Xendit</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_xendit)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Transfer Finlink</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_finlink)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Qty PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_num(lm_qty_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">EWA PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_ewa_ppob)}</td></tr><tr><td style="padding:10px 12px;background:#e0f2fe;border:1px solid #bae6fd;color:#0c4a6e;font-weight:600;font-size:13px;width:48%;">Admin PPOB</td><td style="padding:10px 12px;background:#fff;border:1px solid #e2e8f0;color:#0f172a;font-size:13px;">{format_rp(lm_admin_ppob)}</td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr><td style="width:50%; padding-right:8px; vertical-align:top;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;height:100%;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;border-radius:12px 12px 0 0;">Chart 1 — Finansial</td></tr><tr><td style="padding:12px;text-align:center;vertical-align:middle;"><img src="{chart1_url}" alt="Chart Finansial" style="display:block;width:100%;max-width:100%;height:auto;border-radius:8px;"/></td></tr></table></td><td style="width:50%; padding-left:8px; vertical-align:top;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;height:100%;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;border-radius:12px 12px 0 0;">Chart 2 — Komposisi Profit MTD</td></tr><tr><td style="padding:12px;text-align:center;vertical-align:middle;"><img src="{chart3_url}" alt="Komposisi Profit" style="display:block;width:100%;max-width:100%;height:auto;border-radius:8px;"/></td></tr></table></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Chart 3 — Perbandingan Kuantitas Transaksi</td></tr><tr><td style="padding:12px;"><img src="{chart2_url}" alt="Chart Kuantitas" style="display:block;width:100%;max-width:900px;height:auto;border:1px solid #e2e8f0;border-radius:8px;"/></td></tr></table></td></tr><tr><td style="padding:0 24px 16px 24px;"><table width="100%" style="background:#fff;border:1px solid #e2e8f0;border-radius:12px;"><tr><td style="padding:12px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;color:#0f172a;font-weight:700;font-size:14px;">Insight Singkat</td></tr><tr><td style="padding:12px 16px;color:#0f172a;font-size:13px;line-height:1.6;"><ul style="margin:0;padding-left:18px;"><li>Rata-rata harian MTD Total EWA: <strong>{format_rp(avg_daily_mtd_total)}</strong> vs Bulan Lalu: <strong>{format_rp(avg_daily_lm_total)}</strong> ({diff_total_ewa_mtd_lm:.2f}%).</li><li>Profit hari ini <strong>{format_rp(profit_ewa)}</strong> dibanding rata-rata profit harian MTD <strong>{format_rp(avg_daily_mtd_profit)}</strong> ({diff_profit_daily_mtd:.2f}%).</li><li>Proyeksi 30 hari (run-rate): Total EWA sekitar <strong>{format_rp(runrate_total_ewa)}</strong> dan Profit sekitar <strong>{format_rp(runrate_profit)}</strong>.</li><li>Rasio Profit/EWA hari ini <strong>{margin_pct:.2f}%</strong>; rasio Profit/EWA MTD <strong>{margin_mtd_pct:.2f}%</strong>.</li><li>Kondisi Pending: <strong>{format_num(pending_ewa)} transaksi</strong> dari <strong>{format_num(qty_ewa)}</strong> Qty EWA ({pending_pct:.2f}%).</li></ul></td></tr></table></td></tr><tr><td style="padding:0 24px 24px 24px;"><table width="100%"><tr><td style="height:2px;background:#f97316;line-height:2px;font-size:2px;">&nbsp;</td></tr><tr><td style="padding-top:10px;color:#64748b;font-size:12px;text-align:center;">Disusun otomatis oleh: <b>{st.session_state["username"]}</b> • Waktu: {get_current_time_wib()}</td></tr></table></td></tr></table></td></tr></table></body></html>"""
            st.success("🎉 Laporan Eksekutif berhasil di-generate!")
            st.text_area("📋 Copy Kode HTML Laporan:", value=html_output, height=150)
            st.components.v1.html(html_output, height=1000, scrolling=True)


    # --- MENU 2: PROSES DATA EWA ---
    elif menu_selection == "⚙️ Proses EWA (Mapping)":
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.markdown("### ⚙️ Pemrosesan & Split Data EWA Otomatis")
        st.info("Upload file Data Raw EWA Anda. Sistem otomatis membaca tabel (mengabaikan baris kosong/judul), memisahkan transaksi ke **Xendit** & **Finlink**, serta menyingkirkan status yang belum disetujui.")
        
        uploaded_file = st.file_uploader("Upload File Raw EWA (.xlsx atau .csv)", type=['xlsx', 'csv'])
        
        if uploaded_file:
            try:
                # 1. BACA FILE
                if uploaded_file.name.endswith('.csv'): df_raw = pd.read_csv(uploaded_file)
                else: df_raw = pd.read_excel(uploaded_file)
                
                # 2. BERSIHKAN HEADER BAWAAN
                df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()

                # 3. AUTO-DETECT HEADER (Mata X-Ray untuk melewati baris kosong/judul dokumen)
                if 'STATUS' not in df_raw.columns:
                    header_row_idx = -1
                    # Cek 15 baris pertama, cari kata STATUS
                    for idx, row in df_raw.head(15).iterrows():
                        if any('STATUS' == str(val).strip().upper() for val in row):
                            header_row_idx = idx
                            break
                    
                    if header_row_idx != -1:
                        # Jadikan baris tersebut sebagai header asli
                        new_header = df_raw.iloc[header_row_idx].astype(str).str.strip().str.upper()
                        df_raw = df_raw[header_row_idx + 1:].copy()
                        df_raw.columns = new_header
                        df_raw.reset_index(drop=True, inplace=True)

                # 4. PROSES JIKA HEADER SUDAH BENAR
                if 'STATUS' not in df_raw.columns:
                    st.error("❌ File salah! Tetap tidak ditemukan kolom 'STATUS' pada tabel.")
                else:
                    df_raw['STATUS'] = df_raw['STATUS'].fillna('').astype(str).str.strip().str.upper()
                    
                    df_approved = df_raw[df_raw['STATUS'] == 'DISETUJUI ATASAN']
                    df_skipped = df_raw[df_raw['STATUS'] != 'DISETUJUI ATASAN']
                    
                    xendit_list = []
                    finlink_list = []
                    xendit_kws = ['OCBC', 'BSI', 'ALLO', 'BANTEN', 'BRI', 'BCA', 'JAGO', 'ARTOS', 'DANA', 'SHOPEE', 'NEO', 'SEA', 'KESEJAHTERAAN_EKONOMI', 'GOPAY', 'YUDHA_BHAKTI']
                    
                    for _, row in df_approved.iterrows():
                        bank_code = str(row.get('BANK', '')).strip()
                        is_xendit = any(kw in bank_code.upper() for kw in xendit_kws)
                        
                        if is_xendit:
                            xendit_list.append({
                                'Reference Id': row.get('ID KASBON', ''), 'Amount': row.get('JUMLAH DITRANSFER', ''),
                                'Channel Code': bank_code, 'Account Number': row.get('NO. REKENING', ''),
                                'Account Name': row.get('ATAS NAMA', ''), 'Description': row.get('KETERANGAN', ''),
                                'Email To': row.get('EMAIL', ''), 'Email CC': '', 'Email BCC': ''
                            })
                        else:
                            finlink_info = BANK_MAPPING.get(bank_code, {'finlink_code': 'TIDAK DITEMUKAN', 'finlink_name': 'TIDAK DITEMUKAN - CEK MANUAL'})
                            finlink_list.append({
                                'reference_id': row.get('ID KASBON', ''), 'account_no': row.get('NO. REKENING', ''),
                                'account_name': row.get('ATAS NAMA', ''), 'bank_name': finlink_info['finlink_name'],
                                'bank_code': finlink_info['finlink_code'], 'amount': row.get('JUMLAH DITRANSFER', ''),
                                'status': '', 'description': row.get('KETERANGAN', ''), 'email_to': row.get('EMAIL', ''),
                                'email_cc': '', 'email_bcc': ''
                            })
                            
                    df_x = pd.DataFrame(xendit_list)
                    df_f = pd.DataFrame(finlink_list)
                    
                    st.success(f"✅ Selesai! Berhasil menemukan tabel dan memproses **{len(df_approved)}** data disetujui.")
                    
                    c_x, c_f = st.columns(2)
                    with c_x:
                        st.markdown(f"#### 🟦 Gateway Xendit ({len(df_x)} Data)")
                        if len(df_x) > 0:
                            st.dataframe(df_x.head(3))
                            excel_x = convert_df_to_excel(df_x)
                            st.download_button("⬇️ Download Template Xendit (.xlsx)", excel_x, "Upload_Xendit.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        else: st.info("Tidak ada transaksi untuk Xendit.")
                            
                    with c_f:
                        st.markdown(f"#### 🟧 Gateway Finlink ({len(df_f)} Data)")
                        if len(df_f) > 0:
                            st.dataframe(df_f.head(3))
                            excel_f = convert_df_to_excel(df_f)
                            st.download_button("⬇️ Download Template Finlink (.xlsx)", excel_f, "Upload_Finlink.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                        else: st.info("Tidak ada transaksi untuk Finlink.")
                            
                    if len(df_skipped) > 0:
                        st.markdown("---")
                        st.warning(f"⚠️ **PENGABAIAN DATA:** Terdapat **{len(df_skipped)}** baris data yang di-skip karena status belum disetujui.")
                        st.dataframe(df_skipped[['ID KASBON', 'NAMA', 'JUMLAH DITRANSFER', 'STATUS']])
            
            except Exception as e:
                st.error(f"Terjadi kesalahan teknis saat membaca file: {e}")
        st.markdown("</div>", unsafe_allow_html=True)
