import streamlit as st
import threading
import pandas as pd
import math
import os
import json
from google import genai
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FAYL_KILIDI = threading.Lock()
FAYL_ADI = "Cudo_Klublar_Bazasi.xlsx"

SUTUNLAR = [
    "Klubun yerləşdiyi rayon", "Zalın adı (ümumi)", "Klubun Judo Managerdə adı",
    "Klubun Judo Managerdə kodu", "Klubun Lisenziya kodu", "Sahə (m²)",
    "Zalın ünvanı", "Əlaqə nömrəsi", "Müəllimlərin adı", "Müəllim sayı",
    "Müəllimin yaşı", "Şəxsiyyət vəsiqəsinin FİN kodu", "Qız", "Oğlan",
    "Cəmi idmançı", "Potensial idmançı sayı", "Orta doluluq",
    "Nəticə", "Əlavə not", "Qeyd"
]

RAYONLAR = [
    "Bakı - Binəqədi", "Bakı - Xətai", "Bakı - Xəzər", "Bakı - Qaradağ",
    "Bakı - Nərimanov", "Bakı - Nəsimi", "Bakı - Nizami", "Bakı - Pirallahı",
    "Bakı - Sabunçu", "Bakı - Səbail", "Bakı - Suraxanı", "Bakı - Yasamal",
    "Sumqayıt", "Gəncə", "Mingəçevir", "Şirvan", "Lənkəran", "Şəki", "Yevlax",
    "Abşeron", "Ağcabədi", "Ağdam", "Ağdaş", "Ağstafa", "Ağsu", "Astara",
    "Balakən", "Beyləqan", "Bərdə", "Biləsuvar", "Cəbrayıl", "Cəlilabad",
    "Daşkəsən", "Füzuli", "Gədəbəy", "Goranboy", "Göyçay", "Göygöl",
    "Hacıqabul", "Xaçmaz", "Xızı", "Xocavənd", "İmişli", "İsmayıllı",
    "Kəlbəcər", "Kürdəmir", "Qax", "Qazax", "Qəbələ", "Qobustan", "Quba",
    "Qubadlı", "Qusar", "Laçın", "Lerik", "Masallı", "Neftçala", "Oğuz",
    "Saatlı", "Sabirabad", "Salyan", "Samux", "Siyəzən", "Şabran", "Şamaxı",
    "Şəmkir", "Şuşa", "Tərtər", "Tovuz", "Ucar", "Yardımlı", "Zaqatala",
    "Zəngilan", "Zərdab", "Naxçıvan MR"
]


def hesabla_netice(sahe, qiz, oglan):
    cemi_idmanci = int(qiz) + int(oglan)
    potensial = math.ceil((float(sahe) / 4) * 4)
    if potensial <= 0:
        potensial = 1
    doluluq_faiz = math.ceil((cemi_idmanci / potensial) * 100)

    if doluluq_faiz >= 70:
        netice = "Əla"
    elif doluluq_faiz >= 50:
        netice = "Yaxşı"
    elif doluluq_faiz >= 30:
        netice = "Qənaətbəxş"
    else:
        netice = "Qeyri-qənaətbəxş"

    return cemi_idmanci, potensial, f"{doluluq_faiz}%", netice


def vereqi_dizayn_et(ws):
    basliq_fon = PatternFill(start_color="0F2C59", end_color="0F2C59", fill_type="solid")
    basliq_srift = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    merkez_duzulus = Alignment(horizontal="center", vertical="center", wrap_text=True)
    sol_duzulus = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ince_xett = Side(border_style="thin", color="D9D9D9")
    cercive = Border(left=ince_xett, right=ince_xett, top=ince_xett, bottom=ince_xett)

    reng_ela = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    reng_yaxsi = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    reng_qenaetbexs = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    reng_qeyri = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    ws.row_dimensions[1].height = 34
    for cell in ws[1]:
        cell.fill = basliq_fon
        cell.font = basliq_srift
        cell.alignment = merkez_duzulus
        cell.border = cercive

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 24
        for cell in row:
            cell.border = cercive
            cell.font = Font(name="Calibri", size=11)
            if cell.column in (1, 2, 3, 7, 9, 19, 20):
                cell.alignment = sol_duzulus
            else:
                cell.alignment = merkez_duzulus

            if cell.column == 18:
                val = str(cell.value)
                if val == "Əla":
                    cell.fill = reng_ela
                elif val == "Yaxşı":
                    cell.fill = reng_yaxsi
                elif val == "Qənaətbəxş":
                    cell.fill = reng_qenaetbexs
                elif val == "Qeyri-qənaətbəxş":
                    cell.fill = reng_qeyri

    if ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_uzunluq = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max(max_uzunluq + 5, 14)


def excele_pesekar_yaz(df):
    rayon_sutunu = df["Klubun yerləşdiyi rayon"].astype(str)
    baki_df = df[rayon_sutunu.str.startswith("Bakı")]
    sumqayit_df = df[rayon_sutunu == "Sumqayıt"]
    region_df = df[(~rayon_sutunu.str.startswith("Bakı")) & (rayon_sutunu != "Sumqayıt")]

    with pd.ExcelWriter(FAYL_ADI, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ümumi Baza")
        baki_df.to_excel(writer, index=False, sheet_name="Bakı nüsxəsi")
        sumqayit_df.to_excel(writer, index=False, sheet_name="Sumqayıt nüsxəsi")
        region_df.to_excel(writer, index=False, sheet_name="Region nüsxəsi")

        for sheet_name in ["Ümumi Baza", "Bakı nüsxəsi", "Sumqayıt nüsxəsi", "Region nüsxəsi"]:
            vereqi_dizayn_et(writer.sheets[sheet_name])


st.set_page_config(page_title="Azərbaycan Cüdo Federasiyası | Elektron Baza", page_icon="🥋", layout="wide")

st.markdown("""
<style>
    @import url('https://googleapis.com');

    .stApp {
        background: radial-gradient(circle at 50% 0%, #0D1E3A 0%, #070F1E 55%, #030712 100%) !important;
        color: #F1F5F9 !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }
    h1, h2, h3, h4, p, label, .stMarkdown, div[data-testid="stWidgetLabel"] p {
        color: #F1F5F9 !important;
    }
    .header-banner {
        background: linear-gradient(180deg, rgba(15, 35, 70, 0.85) 0%, rgba(7, 15, 30, 0.98) 100%);
        padding: 42px 36px;
        border-radius: 20px;
        margin-bottom: 28px;
        border: 1px solid rgba(56, 189, 248, 0.25);
        text-align: center;
        box-shadow: 0 20px 45px rgba(0, 0, 0, 0.55);
    }
    .header-badge {
        display: inline-block;
        background-color: rgba(56, 189, 248, 0.14);
        color: #38BDF8 !important;
        border: 1px solid rgba(56, 189, 248, 0.45);
        padding: 5px 16px;
        border-radius: 30px;
        font-size: 12px;
        font-weight: 600 !important;
        letter-spacing: 1px;
        margin-bottom: 14px;
    }
    .header-banner h1 {
        font-family: 'Playfair Display', serif !important;
        color: #38BDF8 !important;
        margin: 0;
        font-size: 38px;
        font-weight: 700;
        letter-spacing: -0.4px;
    }
    .header-banner p {
        color: #94A3B8 !important;
        margin: 12px auto 0 auto;
        font-size: 16px;
        max-width: 680px;
    }
    .info-card, div[data-testid="stForm"] {
        background: #0C182E !important;
        padding: 26px !important;
        border-radius: 16px !important;
        border: 1px solid rgba(56, 189, 248, 0.2) !important;
        box-shadow: 0 14px 32px rgba(0, 0, 0, 0.4) !important;
    }
    input, textarea, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        background-color: #060D1A !important;
        color: #F8FAFC !important;
        border: 1px solid rgba(56, 189, 248, 0.35) !important;
        border-radius: 12px !important;
    }
    input::placeholder, textarea::placeholder {
        color: #64748B !important;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #0C182E !important;
        color: #94A3B8 !important;
        border-radius: 24px !important;
        padding: 10px 24px !important;
        border: 1px solid rgba(56, 189, 248, 0.2) !important;
        margin-right: 8px;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #0284C7 0%, #38BDF8 100%) !important;
        color: #030712 !important;
        font-weight: 700 !important;
    }
    .stTabs [aria-selected="true"] p {
        color: #030712 !important;
    }
    button, .stButton > button, .stDownloadButton > button {
        background: linear-gradient(90deg, #0284C7 0%, #38BDF8 100%) !important;
        color: #030712 !important;
        border: none !important;
        border-radius: 28px !important;
        padding: 10px 24px !important;
        font-weight: 700 !important;
        box-shadow: 0 6px 20px rgba(56, 189, 248, 0.28) !important;
    }
    button p, .stButton > button p, .stDownloadButton > button p {
        color: #030712 !important;
    }
    div[data-testid="metric-container"] {
        background: #0C182E !important;
        border: 1px solid rgba(56, 189, 248, 0.22) !important;
        padding: 22px 24px !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 26px rgba(0, 0, 0, 0.35) !important;
    }
    div[data-testid="stMetricValue"] {
        color: #38BDF8 !important;
        font-family: 'Playfair Display', serif !important;
        font-size: 36px !important;
    }
    div[data-testid="stExpander"] summary {
        background-color: #0C182E !important;
        color: #38BDF8 !important;
        border: 1px solid rgba(56, 189, 248, 0.3) !important;
        border-radius: 14px !important;
    }
</style>
    

""", unsafe_allow_html=True)

st.markdown("""
<div class="header-banner">
    <h1>🥋 AZƏRBAYCAN CÜDO FEDERASİYASI</h1>
    <p>Respublika üzrə Cüdo Klublarının və Zallarının Vahid Elektron Monitorinq Sistemi</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📝 Müəllim Qeydiyyat Portalı", "🔐 Federasiya İdarəetmə Mərkəzi", "❓ Təlimat və Dəstək", "💬 AI Köməkçi"])

with tab1:
    st.info(
        "Hörmətli məşqçi-müəllimlər, klubunuzun bir və ya bir neçə zalı (filialı) varsa, zal sayını seçərək bütün zalları eyni anda bazaya daxil edə bilərsiniz.")

    zal_sayi = st.number_input("Klubun (və ya müəllimin) ümumi zal / filial sayı:", min_value=1, max_value=10, value=1,
                               step=1)

    with st.form("klub_formu", clear_on_submit=True):
        st.markdown("### 👤 Baş Məşqçi və Ümumi Məlumatlar")
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            muellim_adi = st.text_input("Müəllimlərin adı, soyadı *")
            elaqe = st.text_input("Əlaqə nömrəsi (WhatsApp aktiv) *")
        with m_col2:
            muellim_sayi = st.number_input("Ümumi müəllim sayı *", min_value=1, step=1)
            muellim_yasi = st.text_input("Müəllimin yaşı (bir neçədirsə vergüllə yazın) *")
            fin_kod = st.text_input("Baş məşqçinin FİN kodu (7 simvol) *", max_chars=7).upper()

        st.divider()
        zallar_melumat = []

        for i in range(int(zal_sayi)):
            st.markdown(f"### 🏢 Zal #{i + 1} Məlumatları")
            z_col1, z_col2 = st.columns(2)
            with z_col1:
                rayon = st.selectbox(f"Zal #{i + 1} - Yerləşdiyi rayon/şəhər *", RAYONLAR, key=f"rayon_{i}")
                zal_adi = st.text_input(f"Zal #{i + 1} - Zalın adı (ümumi) *", key=f"zal_{i}")
                sahe = st.number_input(f"Zal #{i + 1} - Sahə (m²) *", min_value=1.0, step=1.0, key=f"sahe_{i}")
                unvan = st.text_input(f"Zal #{i + 1} - Zalın dəqiq ünvanı *", key=f"unvan_{i}")
            with z_col2:
                qiz = st.number_input(f"Zal #{i + 1} - Qız idmançı sayı *", min_value=0, step=1, key=f"qiz_{i}")
                oglan = st.number_input(f"Zal #{i + 1} - Oğlan idmançı sayı *", min_value=0, step=1, key=f"oglan_{i}")
                elave_not = st.text_area(f"Zal #{i + 1} - Əlavə not (zalın şəraiti, ehtiyaclar)", key=f"not_{i}")

            zallar_melumat.append({
                "rayon": rayon, "zal_adi": zal_adi, "sahe": sahe,
                "unvan": unvan, "qiz": qiz, "oglan": oglan, "elave_not": elave_not
            })
            st.divider()

        gonder = st.form_submit_button("✅ Məlumatları Federasiya Bazasına Göndər", use_container_width=True,
                                       type="primary")

        if gonder:
            xeta_var = False
            if not muellim_adi or not elaqe or len(fin_kod) != 7:
                xeta_var = True

        for z in zallar_melumat:
            if not z["zal_adi"] or not z["unvan"]:
                xeta_var = True
        if xeta_var:
            st.error("⚠️ Zəhmət olmasa bütün vacib (*) xanaları doldurun və FİN kodun 7 simvol olduğuna əmin olun!")
        else:
            api_key_form = st.secrets.get("GEMINI_API_KEY", "")
            yeni_setirler = []
            for z in zallar_melumat:
                cemi_idmanci, potensial, doluluq_str, netice = hesabla_netice(z["sahe"], z["qiz"], z["oglan"])
                son_not = z["elave_not"]

                if api_key_form and str(son_not).strip() != "":
                    try:
                        c_ai = genai.Client(api_key=api_key_form)
                        r_ai = c_ai.models.generate_content(
                            model="gemini-3.8-flash",
                            contents=f"Aşağıdakı idman zalı qeydini Azərbaycan dilində çox qısa, rəsmi və səliqəli bir cümləyə çevir (yalnız düzəldilmiş mətni yaz): {son_not}"
                        )
                        son_not = r_ai.text.strip()
                    except Exception:
                        pass

                yeni_setirler.append([
                    z["rayon"], z["zal_adi"], "", "", "", z["sahe"], z["unvan"], elaqe,
                    muellim_adi, muellim_sayi, muellim_yasi, fin_kod,
                    z["qiz"], z["oglan"], cemi_idmanci, potensial, doluluq_str,
                    netice, son_not, ""
                ])
                with FAYL_KILIDI:
                    if os.path.exists(FAYL_ADI):
                        df = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza", dtype=str).fillna("")
                        movcud_baza = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza")
                        yenilenmis_baza = pd.concat([movcud_baza, yeni_df], ignore_index=True)
                    else:
                        yenilenmis_baza = yeni_df

                    excele_pesekar_yaz(yenilenmis_baza)
                st.success(f"🎉 Təşəkkürlər, {muellim_adi}! {len(yeni_setirler)} zal üzrə məlumatlar avtomatik olaraq müvafiq nüsxəyə (Bakı / Sumqayıt / Region) əlavə edildi.")

        st.error("❌ Şifrə yanlışdır!")

with tab2:
    sifre = st.text_input("İdarəetmə paneli şifrəsini daxil edin:", type="password")

    if sifre == "judo2026":
        # 1. KÖHNƏ EXCEL BAZASINI SİSTEMƏ YÜKLƏMƏ (İMPORT - V9.2 TAM UYĞUN)
        with st.expander("📂 Mövcud Vahid Excel Bazasını (V9.2) Sistemə Yüklə", expanded=True):
            st.info(
                "V9.2 Excel faylını seçin. Sistem 'Regionlar nüsxəsi', 'Bakı nüsxəsi' və 'Sumqayıt nüsxəsi' vərəqlərini eyni anda oxuyub birləşdirilmiş xanaları avtomatik dolduracaq.")
            yuklenen_fayl = st.file_uploader("V9.2 Excel faylını (.xlsx) seçin:", type=["xlsx"])

            if yuklenen_fayl is not None:
                if st.button("🚀 Bütün Vərəqləri (Region + Bakı + Sumqayıt) Avtomatik İnteqrasiya Et", type="primary"):
                    try:
                        excel_fayl = pd.ExcelFile(yuklenen_fayl)
                        butun_vereqler = []

                        sutun_lugeti = {
                            "Şəhər/rayon": "Klubun yerləşdiyi rayon",
                            "Rayon": "Klubun yerləşdiyi rayon",
                            "Zalın adı (detallı)": "Zalın adı (ümumi)",
                            "Zalın adı": "Zalın adı (ümumi)",
                            "Klubun lisenziya kodu": "Klubun Lisenziya kodu",
                            "Zalın ölçüləri kvm ilə": "Sahə (m²)",
                            "Sahə": "Sahə (m²)",
                            "Zalın Ünvanı": "Zalın ünvanı",
                            "Ünvan": "Zalın ünvanı",
                            "Əlaqə Nömrəsi": "Əlaqə nömrəsi",
                            "Müəllim yaşı": "Müəllimin yaşı",
                            "Müəllimin Şəxsiyyət vəsiqəsinin FİN kodu": "Şəxsiyyət vəsiqəsinin FİN kodu",
                            "QIZ": "Qız",
                            "OĞLAN": "Oğlan",
                            "Potensial": "Potensial idmançı sayı",
                            "Əlavə qeyd": "Əlavə not",
                            "ƏLAVƏ NOT": "Əlavə not"
                        }

                        for v_adi in excel_fayl.sheet_names:
                            v_df = pd.read_excel(yuklenen_fayl, sheet_name=v_adi)
                            # Sütun adlarındakı sətir keçidlərini və artıq boşluqları təmizləyirik
                            v_df.columns = [" ".join(str(c).split()) for c in v_df.columns]
                            v_df = v_df.rename(columns=sutun_lugeti)

                            # Bakı nüsxəsində Rayon sütunu sadəcə rayon adıdırsa, əvvəlinə "Bakı - " artırırıq
                            if "Bakı" in v_adi and "Klubun yerləşdiyi rayon" in v_df.columns:
                                v_df["Klubun yerləşdiyi rayon"] = v_df["Klubun yerləşdiyi rayon"].ffill()
                                v_df["Klubun yerləşdiyi rayon"] = v_df["Klubun yerləşdiyi rayon"].apply(
                                    lambda x: f"Bakı - {str(x).strip()}" if pd.notna(x) and str(
                                        x).strip() != "" and not str(x).strip().startswith("Bakı") else x
                                )

                            # Birləşdirilmiş (Merged) xanaların aşağıya doğru doldurulması
                            doldurulacaq_sutunlar = [
                                "Klubun yerləşdiyi rayon", "Zalın adı (ümumi)",
                                "Klubun Judo Managerdə adı", "Klubun Judo Managerdə kodu",
                                "Klubun Lisenziya kodu", "Sahə (m²)", "Zalın ünvanı"
                            ]
                            for col_name in doldurulacaq_sutunlar:
                                if col_name in v_df.columns:
                                    v_df[col_name] = v_df[col_name].ffill()

                            # Yekun cəm sətirlərini və boş sətirləri təmizləyirik
                            if "Müəllimlərin adı" in v_df.columns:
                                v_df = v_df[v_df["Müəllimlərin adı"].notna() & (
                                            v_df["Müəllimlərin adı"].astype(str).str.strip() != "")]

                            hazir_v = pd.DataFrame(columns=SUTUNLAR)
                            for sutun in SUTUNLAR:
                                if sutun in v_df.columns:
                                    hazir_v[sutun] = v_df[sutun].fillna("")
                                else:
                                    hazir_v[sutun] = ""

                            butun_vereqler.append(hazir_v)

                        yekun_df = pd.concat(butun_vereqler, ignore_index=True)

                        # Əgər bazada əvvəldən məlumat varsa birləşdiririk
                        if os.path.exists(FAYL_ADI):
                            kohne_df = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza")
                            yekun_df = pd.concat([kohne_df, yekun_df], ignore_index=True)

                        excele_pesekar_yaz(yekun_df)
                        st.success(
                            f"🎉 Bütün vərəqlər uğurla oxundu! Cəmi {len(yekun_df)} məşqçi/zal qeydi tam dolu şəkildə sistemə inteqrasiya edildi.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"İnteqrasiya zamanı xəta: {e}")

        # 2. SÜNİ İNTELLEKT (AI) RƏQƏMSAL KOORDİNATOR
        with st.expander("🤖 Süni İntellekt (AI) Rəqəmsal Koordinator (Mesaj Çevirici & Rəhbərlik Analizi)", expanded=True):
            api_key = st.secrets.get("GEMINI_API_KEY", "")
            if not api_key:
                api_key = st.text_input("Google Gemini API Açarını daxil edin:", type="password")
            else:
                st.success("🔑 AI açarı sistem yaddaşından avtomatik qoşuldu.")

            ai_col1, ai_col2 = st.columns(2)
            with ai_col1:
                st.markdown("#### 📲 WhatsApp Mesajını Avtomatik Bazaya Sal")
                wp_mesaj = st.text_area("Məşqçinin WhatsApp mesajını bura yapışdırın:", placeholder="Məs: Salam, mən Əli Vəliyev, yaş 58, FİN 1A2B3C4, tel 0501234567. Gəncə şəhəri OİK zalı, sahə 200 kv, 30 oğlan, 12 qız məşq edir.")
                if st.button("✨ AI ilə Məlumatı Çıxar və Bazaya Əlavə Et"):
                    if not api_key or not wp_mesaj:
                        st.warning("Zəhmət olmasa həm API açarını, həm də mesaj mətnini daxil edin.")
                    else:
                        try:
                            client = genai.Client(api_key=api_key)
                            prompt = f"""
                            Aşağıdakı Azərbaycan dilindəki mətndən cüdo zalı məlumatlarını çıxar və yalnız JSON formatında qaytar (əlavə heç bir mətn yazma):
                            Mətn: {wp_mesaj}
                            JSON açarları:
                            "rayon" (siyahıdan ən uyğunu: {RAYONLAR}),
                            "zal_adi" (mətn),
                            "sahe" (rəqəm, məsələn 150),
                            "unvan" (mətn),
                            "elaqe" (mətn),
                            "muellim_adi" (mətn),
                            "muellim_sayi" (rəqəm, default 1),
                            "muellim_yasi" (mətn),
                            "fin_kod" (7 simvol böyük hərflə, yoxdursa "0000000"),
                            "qiz" (rəqəm, default 0),
                            "oglan" (rəqəm, default 0),
                            "elave_not" (mətn)
                            """
                            response = client.models.generate_content(model="gemini-3.8-flash", contents=prompt)
                            temiz_json = response.text.replace("```json", "").replace("```", "").strip()
                            data = json.loads(temiz_json)

                            c_id, pot, dol, net = hesabla_netice(data["sahe"], data["qiz"], data["oglan"])
                            yeni_ai_df = pd.DataFrame([[
                                data["rayon"], data["zal_adi"], "", "", "", float(data["sahe"]), data["unvan"], data["elaqe"],
                                data["muellim_adi"], int(data["muellim_sayi"]), str(data["muellim_yasi"]), str(data["fin_kod"]).upper(),
                                int(data["qiz"]), int(data["oglan"]), c_id, pot, dol, net, data["elave_not"], "AI ilə əlavə edildi"
                            ]], columns=SUTUNLAR)

                            if os.path.exists(FAYL_ADI):
                                movcud_baza = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza")
                                yenilenmis_baza = pd.concat([movcud_baza, yeni_ai_df], ignore_index=True)
                            else:
                                yenilenmis_baza = yeni_ai_df

                            excele_pesekar_yaz(yenilenmis_baza)
                            st.success(f"✅ AI mesajı uğurla emal etdi və '{data['muellim_adi']}' bazaya əlavə olundu!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"AI emalı zamanı xəta: {e}")

            with ai_col2:
                st.markdown("#### 📊 Rəhbərlik Üçün AI Analitika və Təsnifat")
                b_col1, b_col2 = st.columns(2)
                with b_col1:
                    hesabat_btn = st.button("📑 1 Kliklə Rəsmi Arayış Hazırla", use_container_width=True)
                with b_col2:
                    problem_btn = st.button("🔍 Dublikat və Problemləri Tap", use_container_width=True)

                sual = st.text_input("Və ya bazaya özəl sual verin:", placeholder="Məs: Hansı rayonlarda qız idmançı sayı azdır?")
                sual_btn = st.button("🧠 Sualı Cavablandır")

                if (hesabat_btn or problem_btn or sual_btn):
                    if not api_key or not os.path.exists(FAYL_ADI):
                        st.warning("API açarının qoşulduğuna və bazada məlumat olduğuna əmin olun.")
                    else:
                        baza_df = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza").fillna("")
                        baza_df["Cəmi idmançı"] = pd.to_numeric(baza_df["Cəmi idmançı"], errors="coerce").fillna(0)
                        baza_df["Qız"] = pd.to_numeric(baza_df["Qız"], errors="coerce").fillna(0)
                        baza_df["Oğlan"] = pd.to_numeric(baza_df["Oğlan"], errors="coerce").fillna(0)

                        rayon_xulase = baza_df.groupby("Klubun yerləşdiyi rayon").agg(
                            Zal_Sayi=("Zalın adı (ümumi)", "count"),
                            Qiz=("Qız", "sum"),
                            Oglan=("Oğlan", "sum"),
                            Cemi=("Cəmi idmançı", "sum")
                        ).to_string()

                        qeydler = baza_df[baza_df["Əlavə not"].astype(str).str.strip() != ""][
                            ["Klubun yerləşdiyi rayon", "Zalın adı (ümumi)", "Nəticə", "Əlavə not"]
                        ].head(50).to_string()

                        if hesabat_btn:
                            tapshiriq = "Azərbaycan Cüdo Federasiyası rəhbərliyi üçün rəsmi iclas arayışı hazırla: Bakı, Sumqayıt və Regionlar üzrə zal sayı, qız/oğlan nisbəti və ümumi vəziyyəti qısa maddələrlə yaz."
                        elif problem_btn:
                            tapshiriq = "Bazanı yoxla: 1) Zəif göstəricili rayonları göstər; 2) 'Əlavə not' sütunundakı ehtiyacları (tatami, təmir və s.) qruplaşdır."
                        else:
                            tapshiriq = sual

                        analiz_prompt = f"""
                                        Sən Azərbaycan Cüdo Federasiyasının baş data analitikisən. Aşağıdakı rayonlar üzrə ümumiləşdirilmiş rəqəmlərə və qeydlərə əsasən Azərbaycan dilində dəqiq, rəsmi və strukturlaşdırılmış cavab ver.
                                        Tapşırıq: {tapshiriq}
                                        Ümumi zal sayı: {len(baza_df)}
                                        Rayonlar üzrə statistika:
                                        {rayon_xulase}
                                        Qeydlər və problemlər (ilk 50):
                                        {qeydler}
                                        """

                        client = genai.Client(api_key=api_key)
                        cavab_alindi = False
                        son_xeta = ""
                        for model_adi in ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-2.0-flash"]:
                            try:
                                res = client.models.generate_content(model=model_adi, contents=analiz_prompt)
                                st.markdown(res.text)
                                cavab_alindi = True
                                break
                            except Exception as e:
                                son_xeta = str(e)
                                continue

                        if not cavab_alindi:
                            st.info("ℹ️ Daxili analitik mühərrik sorğunuza uyğun hesabatı hazırladı:")

                            if hesabat_btn:
                                r_ser = baza_df["Klubun yerləşdiyi rayon"].astype(str)
                                baki_df = baza_df[r_ser.str.startswith("Bakı")]
                                sumq_df = baza_df[r_ser == "Sumqayıt"]
                                reg_df = baza_df[(~r_ser.str.startswith("Bakı")) & (r_ser != "Sumqayıt")]

                                cemi_q = int(baza_df["Qız"].sum())
                                cemi_o = int(baza_df["Oğlan"].sum())
                                nisbet = round(cemi_o / cemi_q, 2) if cemi_q > 0 else cemi_o

                                st.markdown(f"""
                                                        ### 📋 Azərbaycan Cüdo Federasiyası — Rəsmi İclas Arayışı
                                                        * **Ümumi qeydiyyatlı zal sayı:** {len(baza_df)} zal
                                                        * **Bölgələr üzrə paylanma:** Bakı — **{len(baki_df)}**, Sumqayıt — **{len(sumq_df)}**, Regionlar — **{len(reg_df)}**
                                                        * **Ümumi idmançı kontingenti:** {cemi_q + cemi_o} nəfər (**Oğlan:** {cemi_o}, **Qız:** {cemi_q})
                                                        * **Respublika üzrə nisbət:** Hər **1** qız idmançıya **{nisbet}** oğlan idmançı düşür.
                                                        """)

                            elif problem_btn:
                                dublikatlar = baza_df[
                                    (baza_df["Şəxsiyyət vəsiqəsinin FİN kodu"].astype(str).str.strip() != "") &
                                    (baza_df["Şəxsiyyət vəsiqəsinin FİN kodu"].astype(str) != "0000000") &
                                    (baza_df.duplicated(subset=["Şəxsiyyət vəsiqəsinin FİN kodu", "Zalın adı (ümumi)"],
                                                        keep=False))
                                    ]
                                zeif_zallar = baza_df[baza_df["Nəticə"].astype(str).str.contains("Qeyri", na=False)]
                                notlu_zallar = baza_df[baza_df["Əlavə not"].astype(str).str.strip() != ""]

                                st.markdown(f"""
                                                        ### 🔍 Baza Audit və Problem Hesabatı
                                                        * **Təkrarlanan (Dublikat) qeydlər:** {len(dublikatlar)} sətir
                                                        * **Qeyri-qənaətbəxş nəticəli zallar:** {len(zeif_zallar)} zal
                                                        * **Xüsusi qeyd/ehtiyac bildirilən zallar:** {len(notlu_zallar)} zal
                                                        """)
                                if len(zeif_zallar) > 0:
                                    st.markdown("**⚠️ Qeyri-qənaətbəxş göstəricili zallar (ilk 15):**")
                                    st.dataframe(zeif_zallar[["Klubun yerləşdiyi rayon", "Zalın adı (ümumi)",
                                                              "Müəllimlərin adı", "Cəmi idmançı", "Orta doluluq"]].head(
                                        15), use_container_width=True)
                                if len(notlu_zallar) > 0:
                                    st.markdown("**📝 Məşqçilərin bildirdiyi əlavə qeydlər və ehtiyaclar (ilk 15):**")
                                    st.dataframe(notlu_zallar[["Klubun yerləşdiyi rayon", "Zalın adı (ümumi)",
                                                               "Əlavə not"]].head(15), use_container_width=True)

                            else:
                                sual_kicik = str(tapshiriq).lower()
                                r_ser = baza_df["Klubun yerləşdiyi rayon"].astype(str)
                                if "region" in sual_kicik:
                                    secim_df = baza_df[(~r_ser.str.startswith("Bakı")) & (r_ser != "Sumqayıt")]
                                    bolge_adi = "Regionlar"
                                elif "bakı" in sual_kicik or "baki" in sual_kicik:
                                    secim_df = baza_df[r_ser.str.startswith("Bakı")]
                                    bolge_adi = "Bakı"
                                elif "sumqayıt" in sual_kicik or "sumqayit" in sual_kicik:
                                    secim_df = baza_df[r_ser == "Sumqayıt"]
                                    bolge_adi = "Sumqayıt"
                                else:
                                    secim_df = baza_df
                                    bolge_adi = "Ümumi Respublika"

                                if "nisbət" in sual_kicik or "nisbet" in sual_kicik:
                                    q_cem = int(secim_df["Qız"].sum())
                                    o_cem = int(secim_df["Oğlan"].sum())
                                    nisbet = round(o_cem / q_cem, 2) if q_cem > 0 else o_cem
                                    st.markdown(
                                        f"**📊 {bolge_adi} üzrə Nisbət:** Oğlan: **{o_cem}**, Qız: **{q_cem}** — Hər **1** qıza **{nisbet}** oğlan düşür (**{o_cem}:{q_cem}**).")
                                else:
                                    rayon_qrup = secim_df.groupby("Klubun yerləşdiyi rayon", as_index=False).agg(
                                        Zal_Sayi=("Zalın adı (ümumi)", "count"),
                                        Qiz=("Qız", "sum"),
                                        Oglan=("Oğlan", "sum"),
                                        Cemi=("Cəmi idmançı", "sum")
                                    )
                                    if "az" in sual_kicik and ("qız" in sual_kicik or "qiz" in sual_kicik):
                                        st.dataframe(rayon_qrup.sort_values(by="Qiz", ascending=True).head(15),
                                                     use_container_width=True)
                                    elif "çox" in sual_kicik or "cox" in sual_kicik or "ən" in sual_kicik:
                                        st.dataframe(rayon_qrup.sort_values(by="Cemi", ascending=False).head(15),
                                                     use_container_width=True)
                                    else:
                                        st.dataframe(rayon_qrup.sort_values(by="Cemi", ascending=False),
                                                     use_container_width=True)
            if os.path.exists(FAYL_ADI):
                df = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza", dtype=str).fillna("")
            else:
                df = pd.DataFrame(columns=SUTUNLAR)
            st.markdown("### ✏️ Bazanı Birbaşa Redaktə Et (Judo Manager, Lisenziya və Qeyd)")
            st.info("Cədvəldəki istənilən xanaya iki dəfə klikləyib məlumatı yazın və aşağıdakı düymə ilə yadda saxlayın. Bütün nüsxələr (Bakı, Sumqayıt, Region) avtomatik yenilənəcək.")

            redakte_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic",
                key="baza_editor"
            )

            if st.button("💾 Bütün Dəyişiklikləri Yadda Saxla", type="primary"):
                for idx, row in redakte_df.iterrows():
                    try:
                        s_val = float(row["Sahə (m²)"])
                        q_val = int(float(row["Qız"]))
                        o_val = int(float(row["Oğlan"]))
                        c_id, pot, dol, net = hesabla_netice(s_val, q_val, o_val)
                        redakte_df.at[idx, "Cəmi idmançı"] = c_id
                        redakte_df.at[idx, "Potensial idmançı sayı"] = pot
                        redakte_df.at[idx, "Orta doluluq"] = dol
                        redakte_df.at[idx, "Nəticə"] = net
                    except Exception:
                        pass

                excele_pesekar_yaz(redakte_df)
                st.success("✅ Dəyişikliklər yadda saxlanıldı və 4 vərəqli Excel bazası yeniləndi!")
                st.rerun()

            st.divider()

            with open(FAYL_ADI, "rb") as file:
                st.download_button(
                    label="📥 Vahid Excel Bazasını Yüklə (Ümumi + Bakı + Sumqayıt + Region)",
                    data=file,
                    file_name="Cudo_Klublar_Bazasi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# --- 3-CÜ BÖLMƏ: TƏLİMAT VƏ DƏSTƏK ---
with tab3:
    st.markdown("### 📋 Məlumatların Daxil Edilməsi Qaydası")
    st.markdown("""
    * **1. Bir neçə zalı olan məşqçilər:** Formanın yuxarısında **"Ümumi zal / filial sayı"** xanasına 2 və ya 3 yazın, hər zal üçün ayrıca bölmə açılacaq.
    * **2. FİN Kod:** Şəxsiyyət vəsiqəsinin aşağı sağ küncündəki kodun **axırdan 1-ci simvolu çıxılmaqla əvvəlki 7 simvolu** (hərf və rəqəmlər) FİN koddur.
    * **3. Zalın sahəsi (m²):** Məşq sahəsini kvadrat metr ilə yazın, doluluq faizi avtomatik hesablanacaq.
    """)
# --- 4-CÜ BÖLMƏ: UNİVERSAL AI SÖHBƏT PƏNCƏRƏSİ ---
with tab4:
    st.markdown("### 💬 Azərbaycan Cüdo Federasiyası — Universal AI Köməkçi")
    st.info("Portaldan istifadə qaydaları, FİN kod, zal ölçüləri, cüdo qaydaları və ya ümumi suallarınızı aşağıdakı söhbət qutusuna yazın.")

    if "ai_mesajlar" not in st.session_state:
        st.session_state.ai_mesajlar = [
            {"rol": "assistant", "metn": "Salam! Mən Azərbaycan Cüdo Federasiyasının rəqəmsal AI köməkçisiyəm. Sizə necə kömək edə bilərəm?"}
        ]

    for m in st.session_state.ai_mesajlar:
        with st.chat_message(m["rol"]):
            st.markdown(m["metn"])

    istifadeci_suali = st.chat_input("Sualınızı bura yazın və Enter sıxın...")

    if istifadeci_suali:
        st.session_state.ai_mesajlar.append({"rol": "user", "metn": istifadeci_suali})
        with st.chat_message("user"):
            st.markdown(istifadeci_suali)

        with st.chat_message("assistant"):
            api_key_chat = st.secrets.get("GEMINI_API_KEY", "")
            baza_xulase_chat = ""
            if os.path.exists(FAYL_ADI):
                try:
                    df_c = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza").fillna("")
                    baza_xulase_chat = f"Bazada hazırda {len(df_c)} zal qeydiyyatdadır."
                except Exception:
                    pass

            chat_prompt = f"""
            Sən Azərbaycan Cüdo Federasiyasının rəsmi universal AI köməkçisisən.
            Vəzifən: Məşqçilərə və istifadəçilərə Azərbaycan dilində səmimi, aydın və peşəkar cavab verməkdir.
            Portal haqqında məlumat:
            - Məşqçilər 'Müəllim Qeydiyyat Portalı' bölməsindən zal sayını seçib məlumatları daxil edirlər.
            - FİN kod şəxsiyyət vəsiqəsinin aşağı sağ küncündəki axırdan 1-ci simvol çıxılmaqla əvvəlki 7 simvoldur.
            - Bir neçə zalı olanlar zal sayını artırmaqla bütün filialları eyni anda yaza bilərlər.
            - Həmçinin cüdo qaydaları, məşq metodikası və ümumi suallara da universal cavab ver.
            {baza_xulase_chat}

            İstifadəçinin sualı: {istifadeci_suali}
            """

            cavab_metni = ""
            xeta_detali = ""
            if api_key_chat:
                for vertex_rejim in [True, False]:
                    if cavab_metni:
                        break
                    try:
                        c_chat = genai.Client(vertexai=vertex_rejim, api_key=api_key_chat)
                        for m_adi in ["gemini-3.8-flash", "gemini-3.5-flash-lite"]:
                            try:
                                r_chat = c_chat.models.generate_content(model=m_adi, contents=chat_prompt)
                                cavab_metni = r_chat.text
                                break
                            except Exception as e:
                                xeta_detali = str(e)
                                continue
                    except Exception as e:
                        xeta_detali = str(e)

            if not cavab_metni:
                cavab_metni = f"⚠️ AI bağlantısı alınmadı (Texniki səbəb: {xeta_detali}). Zəhmət olmasa yuxarıdakı 'Təlimat və Dəstək' bölməsinə baxın."

            st.markdown(cavab_metni)
            st.session_state.ai_mesajlar.append({"rol": "assistant", "metn": cavab_metni})