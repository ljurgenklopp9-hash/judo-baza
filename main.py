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
    .stApp {
        background: linear-gradient(180deg, #E2E8F0 0%, #F8FAFC 40%, #EFF6FF 100%) !important;
        color: #0F172A !important;
    }
    h1, h2, h3, h4, p, label, .stMarkdown, div[data-testid="stWidgetLabel"] p {
        color: #0F172A !important;
        font-weight: 600 !important;
    }
    .header-banner {
        background: linear-gradient(135deg, #06182C 0%, #0F3460 55%, #1E5F74 100%);
        padding: 30px 38px;
        border-radius: 16px;
        margin-bottom: 24px;
        box-shadow: 0 12px 28px rgba(6, 24, 44, 0.22);
        border-left: 10px solid #D97706;
    }
    .header-banner h1 { color: #FFFFFF !important; margin: 0; font-size: 30px; font-weight: 800; }
    .header-banner p { color: #CBD5E1 !important; margin: 8px 0 0 0; font-size: 15px; font-weight: 400 !important; }
    .info-card {
        background: rgba(255, 255, 255, 0.94);
        padding: 22px;
        border-radius: 14px;
        border: 1px solid #CBD5E1;
        border-left: 6px solid #D97706;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
        margin-bottom: 16px;
    }
    div[data-testid="stForm"] {
        background: #FFFFFF !important;
        padding: 26px !important;
        border-radius: 16px !important;
        border: 1px solid #CBD5E1 !important;
        box-shadow: 0 10px 25px rgba(15, 23, 42, 0.07) !important;
    }
    input, textarea, div[data-baseweb="select"] > div, div[data-baseweb="input"] > div {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        border: 1.5px solid #94A3B8 !important;
        border-radius: 10px !important;
    }
    input::placeholder, textarea::placeholder { color: #64748B !important; opacity: 1 !important; }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF !important;
        color: #0F172A !important;
        border-radius: 10px 10px 0 0;
        padding: 12px 24px;
        border: 1px solid #CBD5E1;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #06182C 0%, #0F3460 100%) !important;
        color: #FFFFFF !important;
        border-bottom: 3px solid #D97706 !important;
    }
    .stTabs [aria-selected="true"] p { color: #FFFFFF !important; }
    button, .stButton > button, .stDownloadButton > button {
        background: linear-gradient(90deg, #06182C 0%, #0F3460 100%) !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 12px rgba(6, 24, 44, 0.18) !important;
    }
    button p, .stButton > button p, .stDownloadButton > button p { color: #FFFFFF !important; }
    div[data-testid="metric-container"] {
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1;
        padding: 20px 24px;
        border-radius: 14px;
        box-shadow: 0 6px 16px rgba(15, 23, 42, 0.06);
        border-top: 5px solid #0F3460;
    }
    div[data-testid="stAlert"] p { color: #FFFFFF !important; }
    div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary p {
        background: linear-gradient(90deg, #06182C 0%, #0F3460 100%) !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-banner">
    <h1>🥋 AZƏRBAYCAN CÜDO FEDERASİYASI</h1>
    <p>Respublika üzrə Cüdo Klublarının və Zallarının Vahid Elektron Monitorinq Sistemi</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📝 Müəllim Qeydiyyat Portalı", "🔐 Federasiya İdarəetmə Mərkəzi"])

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

                        for model_adi in ["gemini-2.5-flash", "gemini-2.5-flash-lite"]:
                            try:
                                res = client.models.generate_content(model=model_adi, contents=analiz_prompt)
                                st.markdown(res.text)
                                cavab_alindi = True
                                break
                            except Exception:
                                continue

                        if not cavab_alindi:
                            st.info("ℹ️ Daxili analitik mühərrik sualınıza uyğun nəticəni hazırladı:")
                            sual_kicik = str(tapshiriq).lower()
                            rayon_qrup = baza_df.groupby("Klubun yerləşdiyi rayon", as_index=False).agg(
                                Zal_Sayi=("Zalın adı (ümumi)", "count"),
                                Qiz=("Qız", "sum"),
                                Oglan=("Oğlan", "sum"),
                                Cemi=("Cəmi idmançı", "sum")
                            )

                            if "qız" in sual_kicik or "qiz" in sual_kicik:
                                netice_cedvel = rayon_qrup.sort_values(by="Qiz", ascending=True).head(15)
                                st.markdown("**📉 Qız idmançı sayı ən az olan 15 rayon/bölgə:**")
                                st.dataframe(netice_cedvel, use_container_width=True)
                            elif "oğlan" in sual_kicik or "oglan" in sual_kicik:
                                netice_cedvel = rayon_qrup.sort_values(by="Oglan", ascending=True).head(15)
                                st.markdown("**📉 Oğlan idmançı sayı ən az olan 15 rayon/bölgə:**")
                                st.dataframe(netice_cedvel, use_container_width=True)
                            else:
                                cemi_zal = len(baza_df)
                                cemi_qiz = int(baza_df["Qız"].sum())
                                cemi_oglan = int(baza_df["Oğlan"].sum())
                                zeif_zallar = baza_df[baza_df["Nəticə"].astype(str).str.contains("Qeyri", na=False)]
                                st.markdown(
                                    f"**Ümumi zal:** {cemi_zal} | **Ümumi idmançı:** {cemi_qiz + cemi_oglan} (**Qız:** {cemi_qiz}, **Oğlan:** {cemi_oglan}) | **Zəif zallar:** {len(zeif_zallar)}")
                                st.dataframe(rayon_qrup.sort_values(by="Cemi", ascending=False).head(15),
                                             use_container_width=True)
            if os.path.exists(FAYL_ADI):
                df = pd.read_excel(FAYL_ADI, sheet_name="Ümumi Baza", dtype=str).fillna("")
            else:
                df = pd.DataFrame(columns=SUTUNLAR)
            st.markdown("### 🔍 Sürətli Axtarış və Filtr")
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)
            with f_col1:
                bolge_secimi = st.selectbox("Nüsxə (Bölgə) üzrə:", ["Ümumi Baza", "Bakı nüsxəsi", "Sumqayıt nüsxəsi", "Region nüsxəsi"])
            with f_col2:
                axtaris = st.text_input("Açar söz (Müəllim, Zal, FİN, Nömrə):")
            with f_col3:
                movcud_rayonlar = ["Hamısı"] + sorted(list(set(df["Klubun yerləşdiyi rayon"].astype(str))))
                secilen_rayon = st.selectbox("Konkret rayon üzrə:", movcud_rayonlar)
            with f_col4:
                secilen_netice = st.selectbox("Nəticə üzrə:", ["Hamısı", "Əla", "Yaxşı", "Qənaətbəxş", "Qeyri-qənaətbəxş"])

            filtrlenmis_df = df.copy()
            if bolge_secimi == "Bakı nüsxəsi":
                filtrlenmis_df = filtrlenmis_df[filtrlenmis_df["Klubun yerləşdiyi rayon"].astype(str).str.startswith("Bakı")]
            elif bolge_secimi == "Sumqayıt nüsxəsi":
                filtrlenmis_df = filtrlenmis_df[filtrlenmis_df["Klubun yerləşdiyi rayon"].astype(str) == "Sumqayıt"]
            elif bolge_secimi == "Region nüsxəsi":
                filtrlenmis_df = filtrlenmis_df[(~filtrlenmis_df["Klubun yerləşdiyi rayon"].astype(str).str.startswith("Bakı")) & (filtrlenmis_df["Klubun yerləşdiyi rayon"].astype(str) != "Sumqayıt")]

            if secilen_rayon != "Hamısı":
                filtrlenmis_df = filtrlenmis_df[filtrlenmis_df["Klubun yerləşdiyi rayon"] == secilen_rayon]
            if secilen_netice != "Hamısı":
                filtrlenmis_df = filtrlenmis_df[filtrlenmis_df["Nəticə"] == secilen_netice]
            if axtaris:
                mask = filtrlenmis_df.astype(str).apply(lambda x: x.str.contains(axtaris, case=False, na=False)).any(axis=1)
                filtrlenmis_df = filtrlenmis_df[mask]

            st.caption(f"Tapılan nəticə: {len(filtrlenmis_df)} zal")
            st.dataframe(filtrlenmis_df, use_container_width=True)

            st.divider()

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

