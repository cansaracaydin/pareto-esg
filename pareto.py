import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import time
import io
import os
from fpdf import FPDF

# --- STREAMLIT PAGE CONFIG ---
st.set_page_config(
    page_title="PARETO | Kredi Tahsis & ESG Skorlama",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ADVANCED STYLING (CSS - CORPORATE THEME) ---
st.markdown("""
<style>
    /* Global Background (Dark Theme) */
    .stApp {
        background-color: #0f172a; 
        font-family: 'Inter', sans-serif;
        color: #f8fafc;
    }
    
    /* Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #020617; 
        color: #f8fafc;
        border-right: 1px solid #1e293b;
    }
    section[data-testid="stSidebar"] .stMarkdown h1, 
    section[data-testid="stSidebar"] .stMarkdown h2, 
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #10b981; 
        font-weight: 600;
        font-size: 16px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Premium Metric Cards */
    div[data-testid="metric-container"] {
        background: rgba(30, 41, 59, 0.4);
        border: 1px solid rgba(255, 255, 255, 0.05);
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.2);
        transition: all 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        border-color: rgba(16, 185, 129, 0.3);
    }
    div[data-testid="metric-container"] label {
        color: #94a3b8 !important; 
        font-weight: 500;
        font-size: 13px;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #f8fafc; 
        font-size: 26px;
        font-weight: 600;
    }
    
    /* XAI Text Alerts */
    .xai-success {
        padding: 16px;
        background-color: rgba(6, 78, 59, 0.2); 
        border-left: 3px solid #10b981;
        border-radius: 4px;
        color: #a7f3d0;
        font-size: 14px;
        font-weight: 400;
        line-height: 1.5;
    }
    .xai-danger {
        padding: 16px;
        background-color: rgba(127, 29, 29, 0.2); 
        border-left: 3px solid #ef4444;
        border-radius: 4px;
        color: #fecaca;
        font-size: 14px;
        font-weight: 400;
        line-height: 1.5;
    }
    .xai-warning {
        padding: 16px;
        background-color: rgba(120, 53, 15, 0.2); 
        border-left: 3px solid #f59e0b;
        border-radius: 4px;
        color: #fde68a;
        font-size: 14px;
        font-weight: 400;
        line-height: 1.5;
    }
    
    /* Fix header colors */
    h1, h2, h3, h4 {
        color: #f8fafc !important;
        font-weight: 500;
    }
    
    hr {
        border-top: 1px solid #1e293b;
    }
    
    /* Lineage boxes */
    .lineage-box {
        background-color: rgba(30, 41, 59, 0.5);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 6px;
        padding: 20px;
        margin-bottom: 20px;
    }
    .lineage-title {
        color: #10b981;
        font-weight: 600;
        font-size: 14px;
        margin-bottom: 12px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .lineage-text {
        font-size: 14px;
        color: #cbd5e1;
        line-height: 1.6;
    }
</style>
""", unsafe_allow_html=True)

# --- PDF VE EXCEL OLUŞTURMA FONKSİYONLARI ---
def create_pdf_report(kobi_data):
    pdf = FPDF()
    pdf.add_page()
    
    def tr(text):
        replacements = {'ş':'s', 'Ş':'S', 'ğ':'g', 'Ğ':'G', 'ı':'i', 'İ':'I', 'ç':'c', 'Ç':'C', 'ö':'o', 'Ö':'O', 'ü':'u', 'Ü':'U'}
        for k, v in replacements.items():
            text = str(text).replace(k, v)
        return text
        
    pdf.set_font("Arial", size=14, style='B')
    pdf.cell(200, 10, txt=tr("PARETO Kredi Tahsis ve ESG Degerlendirme Raporu"), ln=1, align='C')
    
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 10, txt=tr("-" * 80), ln=1, align='C')
    
    pdf.set_font("Arial", size=11, style='B')
    pdf.cell(200, 10, txt=tr(f"Musteri No: {kobi_data['KOBI_ID']} | Sektor: {kobi_data['Sektor']}"), ln=1, align='L')
    
    pdf.set_font("Arial", size=10)
    pdf.cell(200, 8, txt=tr(f"Skor Kart Degeri (Basel): {kobi_data['Geleneksel_Kredi_Skoru']}"), ln=1, align='L')
    pdf.cell(200, 8, txt=tr(f"PARETO Dayaniklilik Skoru: {kobi_data['PARETO_ESG_Skoru']:.2f} / 100"), ln=1, align='L')
    pdf.cell(200, 8, txt=tr(f"Beklenen Temerrut Olasiligi (PD): %{kobi_data['PD_Temerrut_Olasiligi_Yuzde']:.2f}"), ln=1, align='L')
    pdf.cell(200, 8, txt=tr(f"Kapsam 1 Emisyon Toplami: {kobi_data['Kapsam_1_Emisyon_Ton']:.1f} tCO2e"), ln=1, align='L')
    pdf.cell(200, 8, txt=tr(f"Nakit Akis Volatilitesi: {kobi_data['Odeme_Sapmasi_Std_Gun']:.1f} Gun"), ln=1, align='L')
    
    pdf.cell(200, 10, txt=tr("-" * 80), ln=1, align='C')
    pdf.set_font("Arial", size=11, style='B')
    pdf.cell(200, 10, txt=tr("VERI DENETIM IZI (AUDIT TRAIL)"), ln=1, align='L')
    pdf.set_font("Arial", size=9)
    pdf.multi_cell(0, 6, txt=tr("- Cevre(E): Islem bazli POS verisi (MCC 5541, 4900). Karbon yogunlugu analitik olarak hesaplanmistir."))
    pdf.multi_cell(0, 6, txt=tr("- Sosyal(S): IKB sistemi SGK dokumleri uzrinden personel sirkulasyonu ve ucret esitsizligi saptanmistir."))
    pdf.multi_cell(0, 6, txt=tr("- Yonetisim(G): Vergi Dairesi API nakit akis sapmasi izlenmistir."))
    
    pdf.cell(200, 10, txt=tr("-" * 80), ln=1, align='C')
    pdf.set_font("Arial", size=11, style='B')
    pdf.cell(200, 10, txt=tr("TAHSIS KOMITESI ONERISI:"), ln=1, align='L')
    
    if kobi_data['PARETO_ESG_Skoru'] >= 65 and kobi_data['Fiziksel_Risk_Kategorisi'] != 'Yüksek Risk':
        decision = "ONAY (PRIME): Firmanin risk-getiri profili banka istahina uygundur. Limit tahsisi onerilir."
    elif kobi_data['Fiziksel_Risk_Kategorisi'] == 'Yüksek Risk' or kobi_data['PARETO_ESG_Skoru'] < 40:
        decision = "RED / TEMINAT ARTIRIMI (EWS): Operasyonel ve fiziksel risk yogunlasmasi tespit edilmistir."
    else:
        decision = "SARTA BAGLI ONAY: Tahsis komitesince detayli risk degerlendirmesi gerekmektedir."
        
    pdf.set_font("Arial", size=10)
    pdf.multi_cell(0, 8, txt=tr(decision))
    
    if kobi_data.get('Taksonomi_Uyumlu_Yatirim', False):
        pdf.set_font("Arial", size=9, style='I')
        pdf.cell(200, 8, txt=tr(" "), ln=1) 
        pdf.multi_cell(0, 6, txt=tr(f"Bilgi: Portfoyde GAR oranini destekleyici ({kobi_data.get('Yesil_Yatirim_Turu', '')}) islemi bulunmaktadir."))
    
    return pdf.output(dest='S').encode('latin1')

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='PARETO_Portfoy_Verisi')
    processed_data = output.getvalue()
    return processed_data


# --- ADIM 1: SENTETİK VERİ ÜRETİMİ (Mock Data) ---
@st.cache_data
def generate_synthetic_data(n=500):
    np.random.seed(42)
    lat_min, lat_max = 36.0, 42.0
    lon_min, lon_max = 26.0, 45.0
    
    data = []
    sectors = ['Tarım', 'İmalat']
    
    for i in range(1, n + 1):
        sector = np.random.choice(sectors, p=[0.4, 0.6])
        lat = np.random.uniform(lat_min, lat_max)
        lon = np.random.uniform(lon_min, lon_max)
        
        fuel_expense = np.random.uniform(10000, 500000) if sector == 'İmalat' else np.random.uniform(5000, 200000)
        elec_expense = np.random.uniform(20000, 800000) if sector == 'İmalat' else np.random.uniform(10000, 300000)
        
        total_emp = np.random.randint(10, 250)
        female_ratio = np.random.uniform(0.1, 0.6)
        female_emp = int(total_emp * female_ratio)
        male_emp = total_emp - female_emp
        pay_gap = np.random.uniform(0.0, 0.4)
        turnover_rate = np.random.uniform(0.05, 0.40)
        
        payment_delay_std = np.random.uniform(1.0, 15.0)
        
        trad_credit_score = int(np.random.normal(1200, 200))
        trad_credit_score = max(0, min(1900, trad_credit_score))
        
        data.append({
            'KOBI_ID': f"MSTR-{i:04d}",
            'Sektor': sector,
            'Enlem': lat,
            'Boylam': lon,
            'MCC_5541_Akaryakit_TL': fuel_expense,
            'MCC_4900_Elektrik_Su_TL': elec_expense,
            'Toplam_Calisan': total_emp,
            'Kadin_Calisan': female_emp,
            'Erkek_Calisan': male_emp,
            'Ucret_Farki_Pay_Gap': pay_gap,
            'Personel_Devir_Hizi': turnover_rate,
            'Odeme_Sapmasi_Std_Gun': payment_delay_std,
            'Geleneksel_Kredi_Skoru': trad_credit_score
        })
        
    return pd.DataFrame(data)

# --- ADIM 2: ARKA PLAN ALGORİTMALARI (Skorlama Motoru) ---
def calculate_esg_scores_for_df(df_input):
    df_scored = df_input.copy()
    
    fuel_emission_factor = 0.0026 
    elec_emission_factor = 0.0045 
    
    df_scored['Kapsam_1_Emisyon_Ton'] = (df_scored['MCC_5541_Akaryakit_TL'] * fuel_emission_factor) + \
                                        (df_scored['MCC_4900_Elektrik_Su_TL'] * elec_emission_factor)
                                        
    df_scored['Calisan_Basina_Emisyon'] = df_scored['Kapsam_1_Emisyon_Ton'] / df_scored['Toplam_Calisan'].replace(0, 1)
                                        
    max_emission = df_scored['Kapsam_1_Emisyon_Ton'].max()
    if pd.isna(max_emission) or max_emission == 0: max_emission = 1
    df_scored['Cevresel_Skor'] = 100 - (df_scored['Kapsam_1_Emisyon_Ton'] / max_emission * 100)
    
    if 'Fiziksel_Risk_Kategorisi' not in df_scored.columns:
        if 'Su_Stresi_Riski' not in df_scored.columns:
            np.random.seed(42)
            df_scored['Su_Stresi_Riski'] = np.random.uniform(0, 100, len(df_scored))
            df_scored['Fay_Hatti_Riski'] = np.random.uniform(0, 100, len(df_scored))
            
        df_scored['Fiziksel_Risk_Skoru'] = (df_scored['Su_Stresi_Riski'] + df_scored['Fay_Hatti_Riski']) / 2
        
        conditions = [
            (df_scored['Fiziksel_Risk_Skoru'] < 33),
            (df_scored['Fiziksel_Risk_Skoru'] < 66),
            (df_scored['Fiziksel_Risk_Skoru'] >= 66)
        ]
        choices = ['Düşük Risk', 'Orta Risk', 'Yüksek Risk']
        df_scored['Fiziksel_Risk_Kategorisi'] = np.select(conditions, choices, default='Bilinmiyor')
    
    df_scored['Kadin_Orani'] = df_scored['Kadin_Calisan'] / df_scored['Toplam_Calisan'].replace(0, 1)
    kadin_skoru = np.where(df_scored['Kadin_Orani'] <= 0.5, df_scored['Kadin_Orani'] / 0.5 * 100, 100)
    pay_gap_skoru = np.clip(100 - (df_scored['Ucret_Farki_Pay_Gap'] * 250), 0, 100)
    turnover_skoru = np.clip(100 - (df_scored['Personel_Devir_Hizi'] * 250), 0, 100)
    
    df_scored['Sosyal_Skor'] = (kadin_skoru * 0.4) + (pay_gap_skoru * 0.4) + (turnover_skoru * 0.2)
    
    df_scored['Yonetisim_Skoru'] = np.clip(100 - (df_scored['Odeme_Sapmasi_Std_Gun'] / 15.0 * 100), 0, 100)
    
    df_scored['PARETO_ESG_Skoru'] = (df_scored['Cevresel_Skor'] * 0.4) + (df_scored['Sosyal_Skor'] * 0.3) + (df_scored['Yonetisim_Skoru'] * 0.3)
    
    base_pd = 20 - (df_scored['Geleneksel_Kredi_Skoru'] / 1900 * 19)
    esg_multiplier = 1.3 - (df_scored['PARETO_ESG_Skoru'] / 100 * 0.6)
    
    df_scored['PD_Temerrut_Olasiligi_Yuzde'] = np.clip(base_pd * esg_multiplier, 0.1, 99.9)
    
    if 'Taksonomi_Uyumlu_Yatirim' not in df_scored.columns:
        np.random.seed(123)
        df_scored['Taksonomi_Uyumlu_Yatirim'] = np.random.choice([True, False], p=[0.15, 0.85], size=len(df_scored))
        df_scored['Yesil_Yatirim_Turu'] = np.where(df_scored['Taksonomi_Uyumlu_Yatirim'], 
                                                   np.random.choice(['Güneş Enerjisi Paneli', 'Atık Su Arıtma', 'Enerji Verimli Makine'], size=len(df_scored)), 
                                                   'Yok')
    
    return df_scored

@st.cache_data
def load_and_score_data():
    raw_df = generate_synthetic_data(500)
    return calculate_esg_scores_for_df(raw_df)

# --- UYGULAMA BAŞLATMA ---
if 'data_loaded' not in st.session_state:
    st.session_state['data_loaded'] = False
elif 'df' in st.session_state and 'Calisan_Basina_Emisyon' not in st.session_state['df'].columns:
    st.session_state['data_loaded'] = False # Eski veri kalmışsa zorla yenile

if not st.session_state['data_loaded']:
    st.markdown("<h1 style='text-align: center; color: #f8fafc; margin-top: 10%; font-size: 40px;'>Kredi Tahsis ve ESG Değerlendirme Modülü</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center; color: #64748b; font-weight: 400;'>Veritabanı Senkronizasyonu Bekleniyor...</h3>", unsafe_allow_html=True)
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    steps = [
        "Sistem başlatılıyor: API Gateway bağlantısı sağlandı.",
        "İşlem Veritabanı sorgulanıyor (MCC Bazlı Analiz)...",
        "İK ve Bordro verileri anonimize edilerek aktarılıyor...",
        "Makine Öğrenmesi (ML) algoritmaları yükleniyor...",
        "Portföy Karbon Emisyonu ve Risk Matrisleri hesaplanıyor...",
        "Kurulum Tamamlandı."
    ]
    
    for i, step in enumerate(steps):
        status_text.markdown(f"<p style='text-align:center; color:#94a3b8; font-size: 14px;'>{step}</p>", unsafe_allow_html=True)
        progress_bar.progress((i + 1) * 16)
        time.sleep(0.4)
        
    st.session_state['df'] = load_and_score_data()
    st.session_state['data_loaded'] = True
    st.rerun()

df = st.session_state['df']

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    
    # Logo Yeri
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("<div style='border: 1px dashed #334155; padding: 20px; text-align: center; color: #64748b; border-radius: 8px; margin-bottom: 20px;'>KURUMSAL LOGO<br><small>(Klasöre 'logo.png' ekleyiniz)</small></div>", unsafe_allow_html=True)
    
    st.markdown("<h2 style='color: #f8fafc; font-size: 20px; font-weight: 600;'>PARETO ANALİTİK</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; font-size: 12px;'>Kurumsal Tahsis Modülü v3.0</p>", unsafe_allow_html=True)
    st.divider()
    
    page_selection = st.radio("SİSTEM MODÜLLERİ", [
        "Portföy Analitiği", 
        "Müşteri Değerlendirme (XAI)",
        "Manuel Müşteri Girişi (Onboarding)",
        "Stres Testi ve Senaryo Analizi",
        "Veri Denetim İzi (Audit Trail)"
    ])
    
    st.divider()
    st.markdown("### Portföy Özeti")
    st.metric("Değerlendirilen Müşteri", f"{len(df)}")
    st.metric("Yüksek Riskli Dağılım", f"{len(df[df['PARETO_ESG_Skoru'] < 40])}", delta="İzleme Listesi", delta_color="inverse")

# --- HEADER ---
st.markdown("<h2 style='margin-bottom: 5px; color: #f8fafc; font-size: 24px;'>Kredi Tahsis ve Portföy Sürdürülebilirlik Yönetimi</h2>", unsafe_allow_html=True)
st.markdown("<p style='color: #94a3b8; font-size: 14px;'>Regülatif standartlara uygun, bağımsız veri kaynaklarıyla beslenen otomatik karar motoru.</p>", unsafe_allow_html=True)
st.divider()

# --- MODULE 1: MAKRO GÖRÜNÜM ---
if page_selection == "Portföy Analitiği":
    
    header_col1, header_col2 = st.columns([4, 1])
    with header_col1:
        st.markdown("### Portföy Konsolidasyon Raporu")
    with header_col2:
        excel_data = to_excel(df)
        st.download_button(label="Veri Setini Dışa Aktar (Excel)",
                           data=excel_data,
                           file_name='PARETO_Portfoy_Analizi.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           use_container_width=True)
    
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    avg_esg = df['PARETO_ESG_Skoru'].mean()
    avg_pd = df['PD_Temerrut_Olasiligi_Yuzde'].mean()
    carbon_intensity = df['Kapsam_1_Emisyon_Ton'].sum() / len(df)
    green_ratio = (df['Taksonomi_Uyumlu_Yatirim'].sum() / len(df)) * 100
    
    kpi1.metric("Ağırlıklı ESG Skoru", f"{avg_esg:.1f} / 100", "+4.2 (Sektörel Sapma)")
    kpi2.metric("Beklenen Temerrüt Olasılığı", f"%{avg_pd:.2f}", "-1.5% (Volatilite Etkisi)", delta_color="inverse")
    kpi3.metric("Ortalama Karbon Yoğunluğu", f"{carbon_intensity:,.0f} tCO2e", "Firma Başına", delta_color="inverse")
    kpi4.metric("Taksonomi Uyumlu Varlık Oranı", f"%{green_ratio:.1f}", "GAR Hedefi: %15.0")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_chart1, col_chart2 = st.columns(2)
    
    with col_chart1:
        st.markdown("#### Risk Yoğunlaşma Matrisi (Kredi vs ESG)")
        fig_scatter = px.scatter(df, x="PARETO_ESG_Skoru", y="PD_Temerrut_Olasiligi_Yuzde",
                                 color="Fiziksel_Risk_Kategorisi",
                                 hover_name="KOBI_ID",
                                 color_discrete_map={'Düşük Risk': '#10b981', 'Orta Risk': '#f59e0b', 'Yüksek Risk': '#ef4444'},
                                 opacity=0.7)
        
        fig_scatter.add_hline(y=10, line_dash="dash", line_color="#334155")
        fig_scatter.add_vline(x=50, line_dash="dash", line_color="#334155")
        
        fig_scatter.update_layout(
            xaxis_title="ESG Performans Skoru",
            yaxis_title="Beklenen Temerrüt Olasılığı (%)",
            yaxis_autorange="reversed", 
            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
            font=dict(color="#94a3b8"),
            margin=dict(t=20, b=20, l=20, r=20), height=380,
            legend_title_text='Fiziksel Risk'
        )
        st.plotly_chart(fig_scatter, use_container_width=True)

    with col_chart2:
        st.markdown("#### Bölgesel Karbon ve İklim Riski Dağılımı")
        color_map = {'Düşük Risk': '#10b981', 'Orta Risk': '#f59e0b', 'Yüksek Risk': '#ef4444'}
        fig_map = px.scatter_mapbox(df, lat="Enlem", lon="Boylam", color="Fiziksel_Risk_Kategorisi",
                                    color_discrete_map=color_map, size="Kapsam_1_Emisyon_Ton",
                                    hover_name="KOBI_ID", zoom=4.2, mapbox_style="carto-positron")
        fig_map.update_layout(margin={"r":0,"t":0,"l":0,"b":0}, height=380, paper_bgcolor='rgba(0,0,0,0)', font=dict(color="#94a3b8"))
        st.plotly_chart(fig_map, use_container_width=True)
        
    st.markdown("### Yeşil Varlık (GAR) Kredilendirme Havuzu")
    df_green = df[df['Taksonomi_Uyumlu_Yatirim'] == True].sort_values(by='PARETO_ESG_Skoru', ascending=False)
    st.dataframe(df_green[['KOBI_ID', 'Sektor', 'PARETO_ESG_Skoru', 'Yesil_Yatirim_Turu', 'PD_Temerrut_Olasiligi_Yuzde']].head(10), use_container_width=True, hide_index=True)


# --- MODULE 2: MİKRO GÖRÜNÜM ---
elif page_selection == "Müşteri Değerlendirme (XAI)":
    
    kobi_list = df['KOBI_ID'].tolist()
    
    col_h1, col_h2 = st.columns([4, 1])
    with col_h1:
        selected_kobi = st.selectbox("Analiz Edilecek Müşteri No:", kobi_list)
    
    kobi_data = df[df['KOBI_ID'] == selected_kobi].iloc[0]
    
    with col_h2:
        st.markdown("<br>", unsafe_allow_html=True)
        pdf_bytes = create_pdf_report(kobi_data)
        st.download_button(label="Müşteri Raporunu Çıkar (PDF)",
                           data=pdf_bytes,
                           file_name=f"PARETO_Rapor_{selected_kobi}.pdf",
                           mime='application/pdf',
                           use_container_width=True)
                           
    st.markdown(f"#### Müşteri Profili: {selected_kobi} | Sektör: {kobi_data['Sektor']}")
    
    c1, c2, c3, c4 = st.columns(4)
    
    c1.metric("Skor Kart Değeri (Klasik)", f"{kobi_data['Geleneksel_Kredi_Skoru']:.0f} / 1900")
    c2.metric("PARETO ESG Skoru", f"{kobi_data['PARETO_ESG_Skoru']:.1f} / 100")
    c3.metric("Revize Temerrüt Olasılığı", f"%{kobi_data['PD_Temerrut_Olasiligi_Yuzde']:.2f}")
    c4.metric("Çalışan Başına Emisyon", f"{kobi_data['Calisan_Basina_Emisyon']:.1f} tCO2", "Karbon Yoğunluğu", delta_color="inverse")

    st.markdown("<br>", unsafe_allow_html=True)
    
    col_radar, col_xai = st.columns([2, 3])
    
    with col_radar:
        st.markdown("#### Sektörel Kıyaslama (Benchmarking)")
        sector_avg = df[df['Sektor'] == kobi_data['Sektor']].mean(numeric_only=True)
        
        categories = ['Çevresel Skoru', 'Sosyal Skoru', 'Yönetişim Skoru']
        fig_radar = go.Figure()
        
        fig_radar.add_trace(go.Scatterpolar(
            r=[sector_avg['Cevresel_Skor'], sector_avg['Sosyal_Skor'], sector_avg['Yonetisim_Skoru']],
            theta=categories, fill='toself', name='Sektör Ortalaması',
            line_color='rgba(148, 163, 184, 0.8)', fillcolor='rgba(148, 163, 184, 0.1)'
        ))
        fig_radar.add_trace(go.Scatterpolar(
            r=[kobi_data['Cevresel_Skor'], kobi_data['Sosyal_Skor'], kobi_data['Yonetisim_Skoru']],
            theta=categories, fill='toself', name=f'Müşteri',
            line_color='rgba(16, 185, 129, 1.0)', fillcolor='rgba(16, 185, 129, 0.4)'
        ))
        
        fig_radar.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 100], tickfont=dict(color="#94a3b8")),
                angularaxis=dict(tickfont=dict(color="#f8fafc"))
            ),
            showlegend=True, margin=dict(t=30, b=30, l=30, r=30), height=350,
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#94a3b8"),
            legend=dict(font=dict(color="#94a3b8"), orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    with col_xai:
        st.markdown("#### Karar Çarpanları (SHAP Analizi)")
        base_val = 50 
        env_impact = (kobi_data['Cevresel_Skor'] - 50) * 0.4
        soc_impact = (kobi_data['Sosyal_Skor'] - 50) * 0.3
        gov_impact = (kobi_data['Yonetisim_Skoru'] - 50) * 0.3
        
        fig_waterfall = go.Figure(go.Waterfall(
            orientation="v", measure=["absolute", "relative", "relative", "relative", "total"],
            x=["Baz Puan", "Çevresel Faktör", "Sosyal Faktör", "Yönetişim Faktörü", "Nihai Skor"],
            textposition="outside",
            text=[f"{base_val}", f"{env_impact:+.1f}", f"{soc_impact:+.1f}", f"{gov_impact:+.1f}", f"{kobi_data['PARETO_ESG_Skoru']:.1f}"],
            y=[base_val, env_impact, soc_impact, gov_impact, kobi_data['PARETO_ESG_Skoru']],
            connector={"line":{"color":"#334155"}},
            decreasing={"marker":{"color":"#ef4444"}},
            increasing={"marker":{"color":"#10b981"}},
            totals={"marker":{"color":"#3b82f6"}}
        ))
        fig_waterfall.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=350, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8"))
        st.plotly_chart(fig_waterfall, use_container_width=True)

    st.markdown("---")
    
    st.markdown("### Tahsis Komitesi Karar Önerisi")
    if kobi_data['PARETO_ESG_Skoru'] >= 65 and kobi_data['Fiziksel_Risk_Kategorisi'] != 'Yüksek Risk':
        st.markdown(f"<div class='xai-success'><strong>ONAY (PRIME):</strong> Firmanın risk-getiri profili banka iştahına uygundur. Finansal metriklerde stabilite ve düşük karbon yoğunluğu tespit edilmiştir. Limit tahsisi önerilir.</div>", unsafe_allow_html=True)
    elif kobi_data['Fiziksel_Risk_Kategorisi'] == 'Yüksek Risk' or kobi_data['PARETO_ESG_Skoru'] < 40:
        reason = "lokasyonel iklim/afet riski" if kobi_data['Fiziksel_Risk_Kategorisi'] == 'Yüksek Risk' else "operasyonel zayıflık ve yüksek karbon maliyetleri"
        st.markdown(f"<div class='xai-danger'><strong>RED / TEMİNAT ARTIRIMI (EWS):</strong> Tahsis kriterleri karşılanmamaktadır. Firmanın operasyonları {reason} sebebiyle regülatif dışlanma riski taşımaktadır. İşlem reddi veya ilave teminat tesisi önerilir.</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='xai-warning'><strong>ŞARTA BAĞLI ONAY:</strong> Firmanın temel finansal göstergeleri makul olmakla birlikte, nakit akış volatilitesi {kobi_data['Odeme_Sapmasi_Std_Gun']:.1f} gün sapma göstermektedir. Tahsis komitesince ilave şartlı değerlendirme gerekmektedir.</div>", unsafe_allow_html=True)
        
    if kobi_data['Taksonomi_Uyumlu_Yatirim']:
        st.info(f"Portföy Yönetimi Notu: Firmada ({kobi_data['Yesil_Yatirim_Turu']}) yatırımı tespit edilmiştir. İlgili firmaya GAR oranını destekleyici Yeşil Finansman olanakları sunulabilir.")

# --- MODULE 2.5: MANUEL MÜŞTERİ GİRİŞİ (ONBOARDING) ---
elif page_selection == "Manuel Müşteri Girişi (Onboarding)":
    st.markdown("### Yeni Müşteri Veri Girişi (Onboarding)")
    st.markdown("Bankacılık sistemlerinde geçmişi bulunmayan yeni müşteriler için manuel tahsis parametreleri girişi ve anlık skorlama simülasyonu.")
    st.divider()
    
    with st.form("onboarding_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            kobi_id = st.text_input("Müşteri No (Örn: YENI-001)", value="YENI-001")
            sektor = st.selectbox("Sektör", ["İmalat", "Tarım"])
            geleneksel_skor = st.number_input("Skor Kart Değeri (Basel)", min_value=0, max_value=1900, value=1200)
            calisan = st.number_input("Toplam Çalışan Sayısı", min_value=1, value=50)
            kadin_calisan = st.number_input("Kadın Çalışan Sayısı", min_value=0, value=15)
        with col2:
            akaryakit = st.number_input("Aylık Akaryakıt Gideri (TL)", min_value=0.0, value=50000.0)
            elektrik = st.number_input("Aylık Elektrik Gideri (TL)", min_value=0.0, value=120000.0)
            pay_gap = st.slider("Cinsiyet Ücret Eşitsizliği Endeksi", 0.0, 1.0, 0.15)
        with col3:
            turnover = st.slider("Aylık Personel Devir Hızı", 0.0, 1.0, 0.10)
            gecikme = st.number_input("Vergi/SGK Gecikme Volatilitesi (Gün)", min_value=0.0, value=5.0)
            fiziksel_risk = st.selectbox("Lokasyonel Fiziksel Risk", ["Düşük Risk", "Orta Risk", "Yüksek Risk"])
            yesil_mi = st.checkbox("Taksonomi Uyumlu Yatırım Var Mı?")
            yesil_tur = st.selectbox("Yatırım Türü", ["Yok", "Güneş Enerjisi", "Atık Su Arıtma", "Enerji Verimliliği"])
            
        submitted = st.form_submit_button("Sisteme Al ve Skorla")
        
    if submitted:
        st.markdown("---")
        st.markdown(f"#### Skorlama Sonucu: {kobi_id}")
        new_data = {
            'KOBI_ID': kobi_id,
            'Sektor': sektor,
            'Enlem': 39.0, 'Boylam': 35.0, 
            'MCC_5541_Akaryakit_TL': akaryakit,
            'MCC_4900_Elektrik_Su_TL': elektrik,
            'Toplam_Calisan': calisan,
            'Kadin_Calisan': kadin_calisan,
            'Erkek_Calisan': calisan - kadin_calisan,
            'Ucret_Farki_Pay_Gap': pay_gap,
            'Personel_Devir_Hizi': turnover,
            'Odeme_Sapmasi_Std_Gun': gecikme,
            'Geleneksel_Kredi_Skoru': geleneksel_skor,
            'Taksonomi_Uyumlu_Yatirim': yesil_mi,
            'Yesil_Yatirim_Turu': yesil_tur if yesil_mi else 'Yok',
            'Fiziksel_Risk_Kategorisi': fiziksel_risk
        }
        df_new = pd.DataFrame([new_data])
        df_new_scored = calculate_esg_scores_for_df(df_new).iloc[0]
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Hesaplanan ESG Skoru", f"{df_new_scored['PARETO_ESG_Skoru']:.1f} / 100")
        c2.metric("Beklenen Temerrüt Olasılığı", f"%{df_new_scored['PD_Temerrut_Olasiligi_Yuzde']:.2f}")
        c3.metric("Kapsam 1 Emisyon", f"{df_new_scored['Kapsam_1_Emisyon_Ton']:.1f} tCO2")
        
        if df_new_scored['PARETO_ESG_Skoru'] >= 65 and df_new_scored['Fiziksel_Risk_Kategorisi'] != 'Yüksek Risk':
            st.markdown(f"<div class='xai-success'><strong>ONAY:</strong> Yeni müşteri profili kredi politikalarına uygundur.</div>", unsafe_allow_html=True)
        else:
            st.markdown(f"<div class='xai-warning'><strong>DİKKAT:</strong> Müşterinin skorları yeterli limite ulaşmamıştır. Özel komite onayı gereklidir.</div>", unsafe_allow_html=True)


# --- MODULE 3: SENARYO VE STRES TESTİ ---
elif page_selection == "Stres Testi ve Senaryo Analizi":
    st.markdown("### Etkileşimli Stres Testi (What-If Analizi)")
    st.markdown("Müşteri davranışlarındaki potansiyel değişimlerin, portföy riski ve 5 yıllık temerrüt olasılığı (PD) üzerindeki marjinal etkisini simüle ediniz.")
    
    selected_kobi = st.selectbox("Simülasyon İçin Müşteri Seçiniz:", df['KOBI_ID'].tolist())
    kobi_data = df[df['KOBI_ID'] == selected_kobi].iloc[0].to_dict() 
    
    col_controls, col_results = st.columns([1, 1])
    
    with col_controls:
        st.markdown("#### Bağımsız Değişkenler (Senaryo Girdileri)")
        new_fuel = st.slider("Aylık Akaryakıt Tüketim Maliyeti (TL)", 
                             min_value=0.0, max_value=float(kobi_data['MCC_5541_Akaryakit_TL'])*2, 
                             value=float(kobi_data['MCC_5541_Akaryakit_TL']), step=1000.0)
                             
        new_elec = st.slider("Aylık Elektrik Tüketim Maliyeti (TL)", 
                             min_value=0.0, max_value=float(kobi_data['MCC_4900_Elektrik_Su_TL'])*2, 
                             value=float(kobi_data['MCC_4900_Elektrik_Su_TL']), step=1000.0)
                             
        new_gap = st.slider("Cinsiyet Ücret Eşitsizliği Endeksi", 
                            min_value=0.0, max_value=1.0, 
                            value=float(kobi_data['Ucret_Farki_Pay_Gap']), step=0.05)
                            
        new_delay = st.slider("Vergi/SGK Nakit Çıkışı Gecikme Volatilitesi (Gün)", 
                              min_value=0.0, max_value=30.0, 
                              value=float(kobi_data['Odeme_Sapmasi_Std_Gun']), step=1.0)
                              
    with col_results:
        st.markdown("#### Simülasyon Çıktıları")
        kobi_data['MCC_5541_Akaryakit_TL'] = new_fuel
        kobi_data['MCC_4900_Elektrik_Su_TL'] = new_elec
        kobi_data['Ucret_Farki_Pay_Gap'] = new_gap
        kobi_data['Odeme_Sapmasi_Std_Gun'] = new_delay
        
        df_sim = pd.DataFrame([kobi_data])
        df_sim_scored = calculate_esg_scores_for_df(df_sim).iloc[0]
        
        old_esg = df[df['KOBI_ID'] == selected_kobi].iloc[0]['PARETO_ESG_Skoru']
        new_esg = df_sim_scored['PARETO_ESG_Skoru']
        
        old_pd = df[df['KOBI_ID'] == selected_kobi].iloc[0]['PD_Temerrut_Olasiligi_Yuzde']
        new_pd = df_sim_scored['PD_Temerrut_Olasiligi_Yuzde']
        
        c1, c2 = st.columns(2)
        c1.metric("Simüle Edilmiş ESG Skoru", f"{new_esg:.1f}", f"{new_esg - old_esg:+.1f} Marjinal Etki")
        c2.metric("Simüle Edilmiş Temerrüt Olasılığı", f"%{new_pd:.2f}", f"{new_pd - old_pd:+.2f}% Risk Değişimi", delta_color="inverse")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        if new_pd < old_pd:
            st.markdown("<div class='xai-success'><strong>Risk İyileşmesi Saptandı:</strong> Parametrik varsayımlar firmanın temerrüt riskini düşürmektedir.</div>", unsafe_allow_html=True)
        elif new_pd > old_pd:
            st.markdown("<div class='xai-danger'><strong>Stres Testi Başarısız:</strong> Parametrik değişimler finansal kırılganlığı artırmaktadır.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='xai-warning'><strong>Etki Nötr:</strong> Değişkenlerin mevcut kredi kalifikasyonu üzerinde belirgin bir saptırıcı etkisi bulunmamaktadır.</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### 5 Yıllık Temerrüt Olasılığı (PD) Projeksiyonu")
    
    years = [2026, 2027, 2028, 2029, 2030]
    base_pd_trend = [old_pd + (i * 0.2) for i in range(5)]
    drift = 0.05 if new_pd < old_pd else (0.5 if new_pd > old_pd else 0.2)
    sim_pd_trend = [new_pd + (i * drift) for i in range(5)]
    
    df_proj = pd.DataFrame({
        "Yıl": years * 2,
        "PD (%)": base_pd_trend + sim_pd_trend,
        "Senaryo Tipi": ["Mevcut Durum (Base)"]*5 + ["Stres Senaryosu"]*5
    })
    
    fig_proj = px.line(df_proj, x="Yıl", y="PD (%)", color="Senaryo Tipi", markers=True,
                       color_discrete_map={"Mevcut Durum (Base)": "#94a3b8", "Stres Senaryosu": "#10b981" if new_pd < old_pd else "#ef4444"})
                       
    fig_proj.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", 
        font=dict(color="#94a3b8"), height=350, margin=dict(t=10, b=10),
        xaxis=dict(tickmode='linear', dtick=1)
    )
    st.plotly_chart(fig_proj, use_container_width=True)


# --- MODULE 4: VERİ DENETİM İZİ (AUDIT TRAIL) ---
elif page_selection == "Veri Denetim İzi (Audit Trail)":
    st.markdown("### Veri Denetim İzi (Dinâmik Şeffaflık Modülü)")
    st.markdown("Algoritmik kararların banka iç sistemlerinden gelen ham verilerle nasıl hesaplandığını adım adım denetleyin (Audit Trail). Bu ekran, Black-box yapısını kırarak %100 regülatif şeffaflık sağlar.")
    
    selected_kobi = st.selectbox("Matematiksel Dökümü Çıkarılacak Müşteri:", df['KOBI_ID'].tolist())
    kobi_data = df[df['KOBI_ID'] == selected_kobi].iloc[0]
    
    st.divider()
    
    col_e, col_s, col_g = st.columns(3)
    
    with col_e:
        st.markdown("<div class='lineage-box'>", unsafe_allow_html=True)
        st.markdown("<div class='lineage-title'>Çevresel (E) Karar İzi</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='lineage-text'>
        <strong>API Kaynağı:</strong> CORE_BANKING_POS_DB<br><br>
        <strong>Süzülen Ham Veri (Aylık):</strong><br>
        • MCC 5541 (Akaryakıt): {kobi_data['MCC_5541_Akaryakit_TL']:,.0f} TL<br>
        • MCC 4900 (Elektrik): {kobi_data['MCC_4900_Elektrik_Su_TL']:,.0f} TL<br><br>
        <strong>Dönüşüm Çarpanı (Katsayı):</strong><br>
        (Akaryakıt x 0.0026) + (Elektrik x 0.0045)<br><br>
        <strong>Kapsam 1 Emisyonu:</strong><br>
        = {kobi_data['Kapsam_1_Emisyon_Ton']:,.1f} tCO2e<br><br>
        <strong>Nihai E Skoru:</strong> <span style='color:#10b981; font-weight:bold; font-size:18px;'>{kobi_data['Cevresel_Skor']:.1f} / 100</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col_s:
        st.markdown("<div class='lineage-box'>", unsafe_allow_html=True)
        st.markdown("<div class='lineage-title'>Sosyal (S) Karar İzi</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='lineage-text'>
        <strong>API Kaynağı:</strong> SGK_PAYROLL_API_v2<br><br>
        <strong>Süzülen Ham Veri (Aylık):</strong><br>
        • Aktif Çalışan Sayısı: {kobi_data['Toplam_Calisan']}<br>
        • Kadın Çalışan Oranı: %{(kobi_data['Kadin_Calisan']/kobi_data['Toplam_Calisan'])*100:.0f}<br>
        • Ölçülen Ücret Farkı (Gap): %{kobi_data['Ucret_Farki_Pay_Gap']*100:.0f}<br><br>
        <strong>Dönüşüm Çarpanı (Ağırlık):</strong><br>
        (Kadın Oranı Skoru x 0.4) + (Gap Skoru x 0.4) + (Devir Hızı x 0.2)<br><br>
        <strong>Nihai S Skoru:</strong> <span style='color:#10b981; font-weight:bold; font-size:18px;'>{kobi_data['Sosyal_Skor']:.1f} / 100</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col_g:
        st.markdown("<div class='lineage-box'>", unsafe_allow_html=True)
        st.markdown("<div class='lineage-title'>Yönetişim (G) Karar İzi</div>", unsafe_allow_html=True)
        st.markdown(f"""
        <div class='lineage-text'>
        <strong>API Kaynağı:</strong> GIB_VERGI_API & CORE_NAKIT<br><br>
        <strong>Süzülen Ham Veri (6 Aylık):</strong><br>
        • Gecikmeli Ödeme Sayısı İzleniyor<br>
        • Nakit Akışı Sapması: {kobi_data['Odeme_Sapmasi_Std_Gun']:.1f} Gün<br><br>
        <strong>Dönüşüm Çarpanı (Ceza Modeli):</strong><br>
        100 - (Sapma Gün / 15.0 x 100)<br>
        (15 günü aşan gecikme volatilitesi cezalandırılır)<br><br>
        <strong>Nihai G Skoru:</strong> <span style='color:#10b981; font-weight:bold; font-size:18px;'>{kobi_data['Yonetisim_Skoru']:.1f} / 100</span>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #64748b; font-size: 12px;'>PARETO Kurumsal Analitik Modülü © 2026</p>", unsafe_allow_html=True)
