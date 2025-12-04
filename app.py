import streamlit as st
import pandas as pd
import plotly.express as px
import os
import zipfile
import tempfile
import shutil
import io

# ---------------------------------------------------------
# 1. SAYFA AYARLARI
# ---------------------------------------------------------
st.set_page_config(page_title="Takas, Virman & Hacim Analizi", page_icon="📦", layout="wide")

st.markdown("""
    <style>
    .main { padding: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Klasör Bazlı Takas & Hacim Analizi (ZIP Yükleme)")
st.info("""
**Nasıl Kullanılır?**
1. Bilgisayarınızdaki **'takas'**, **'akd'** ve **'hacim'** klasörlerine sağ tıklayıp **'ZIP dosyasına sıkıştır'** deyin.
2. Oluşan ZIP dosyalarını aşağıya yükleyin. Sistem klasör yapısını (Yıl/Ay) otomatik tanıyacaktır.
""")
st.markdown("---")

# ---------------------------------------------------------
# 2. YARDIMCI FONKSİYONLAR
# ---------------------------------------------------------

def clean_takas_value(val):
    """Excel'den gelen veriyi sayıya çevirir."""
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return val
    
    val_str = str(val).strip()
    val_str = val_str.replace(".", "")
    val_str = val_str.replace(",", ".")
    try:
        return float(val_str)
    except:
        return 0

def extract_zip_and_get_files(uploaded_zip, file_type="takas"):
    """Yüklenen ZIP dosyasını geçici klasöre çıkarır ve kronolojik sıralar."""
    file_list = []
    temp_dir = tempfile.mkdtemp()
    
    try:
        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                if not file.endswith(".xlsx") or file.startswith("~$") or "__MACOSX" in root:
                    continue
                
                rel_path = os.path.relpath(root, temp_dir)
                path_parts = rel_path.split(os.sep)
                
                year = 2024
                month = 1
                
                for part in path_parts:
                    if part.isdigit():
                        val = int(part)
                        if val > 2000:
                            year = val
                        elif 1 <= val <= 12:
                            month = val
                
                full_path = os.path.join(root, file)
                
                try:
                    name_parts = file.replace(".xlsx", "").split()
                    
                    if file_type == "takas":
                        day = int(name_parts[0])
                        sort_key = (year, month, day)
                        display_date = f"{day}.{month}.{year}"
                        
                    elif file_type in ["akd", "hacim"]:
                        day_range = name_parts[0]
                        start_day = int(day_range.split("-")[0])
                        sort_key = (year, month, start_day)
                        display_date = f"{day_range}.{month}.{year}"
                    
                    file_list.append({
                        "path": full_path,
                        "filename": file,
                        "sort_key": sort_key,
                        "display": display_date
                    })
                except Exception:
                    continue

        file_list.sort(key=lambda x: x["sort_key"])
        return file_list, temp_dir

    except zipfile.BadZipFile:
        st.error("Yüklenen dosya geçerli bir ZIP dosyası değil.")
        return [], None
    except Exception as e:
        st.error(f"ZIP açılırken hata oluştu: {e}")
        return [], None

def process_hacim_files(hacim_files):
    """Hacim dosyalarını işler ve haftalık yüzdeleri hesaplar."""
    all_data = []
    
    for file_info in hacim_files:
        df = pd.read_excel(file_info["path"])
        
        # Kurum bazlı toplam
        grouped = df.groupby("Kurum", as_index=False)["Toplam"].sum()
        
        # Haftalık grand total
        grand_total = grouped["Toplam"].sum()
        
        # Yüzde hesaplama
        grouped["Yüzde"] = (grouped["Toplam"] / grand_total * 100).round(2)
        
        grouped.rename(columns={"Toplam": "Haftalık Kurum Toplam"}, inplace=True)
        grouped["Haftalık Toplam"] = grand_total
        grouped["Hafta"] = file_info["display"]
        
        # ALL satırı ekle
        all_row = pd.DataFrame({
            "Kurum": ["ALL"],
            "Haftalık Kurum Toplam": [grand_total],
            "Yüzde": [100.0],
            "Haftalık Toplam": [grand_total],
            "Hafta": [file_info["display"]]
        })
        
        grouped = pd.concat([grouped, all_row], ignore_index=True)
        all_data.append(grouped)
    
    final_df = pd.concat(all_data, ignore_index=True)
    return final_df

# ---------------------------------------------------------
# 3. SIDEBAR (DOSYA YÜKLEME)
# ---------------------------------------------------------
with st.sidebar:
    st.header("📂 Dosya Yükleme")
    
    st.subheader("1️⃣ Takas Klasörü (ZIP)")
    takas_zip = st.file_uploader("Takas.zip yükleyin", type="zip", key="takas")
    
    st.subheader("2️⃣ AKD Klasörü (ZIP)")
    akd_zip = st.file_uploader("AKD.zip yükleyin", type="zip", key="akd")
    
    st.subheader("3️⃣ Hacim Klasörü (ZIP) - Opsiyonel")
    st.caption("İsterseniz Hacim analizi için de yükleyin")
    hacim_zip = st.file_uploader("Hacim.zip yükleyin (opsiyonel)", type="zip", key="hacim")
    
    st.markdown("---")
    process_button = st.button("🚀 Analizi Başlat", type="primary")

# ---------------------------------------------------------
# 4. İŞLEM MANTIĞI
# ---------------------------------------------------------
if process_button:
    if not takas_zip or not akd_zip:
        st.error("❌ Lütfen en az Takas ve AKD ZIP dosyalarını yükleyin.")
    else:
        takas_temp_dir = None
        akd_temp_dir = None
        hacim_temp_dir = None
        
        with st.spinner("📦 ZIP dosyaları açılıyor ve analiz ediliyor..."):
            try:
                # 1. Dosyaları Çıkar
                takas_files, takas_temp_dir = extract_zip_and_get_files(takas_zip, "takas")
                akd_files, akd_temp_dir = extract_zip_and_get_files(akd_zip, "akd")
                
                # Hacim opsiyonel
                hacim_files = []
                if hacim_zip:
                    hacim_files, hacim_temp_dir = extract_zip_and_get_files(hacim_zip, "hacim")
                
                if not takas_files or not akd_files:
                    st.error("ZIP içeriğinde uygun Excel dosyaları bulunamadı.")
                else:
                    success_msg = f"✅ {len(takas_files)} Takas ve {len(akd_files)} AKD dosyası bulundu."
                    if hacim_files:
                        success_msg += f" + {len(hacim_files)} Hacim dosyası"
                    st.success(success_msg)
                    
                    # 2. Takas Farklarını Hesapla
                    diff_list = []
                    for i in range(1, len(takas_files)):
                        prev = takas_files[i - 1]
                        curr = takas_files[i]
                        
                        df_prev = pd.read_excel(prev["path"])
                        df_curr = pd.read_excel(curr["path"])
                        
                        df_prev["Takas"] = df_prev["Takas"].apply(clean_takas_value)
                        df_curr["Takas"] = df_curr["Takas"].apply(clean_takas_value)
                        
                        df_merged = pd.merge(
                            df_curr, df_prev, on="Kurum", 
                            suffixes=("_current", "_previous"), how="outer"
                        )
                        
                        df_merged["Takas_Diff"] = df_merged["Takas_current"].fillna(0) - df_merged["Takas_previous"].fillna(0)
                        df_merged["Week"] = f"{prev['display']} - {curr['display']}"
                        
                        diff_list.append(df_merged)
                    
                    if diff_list:
                        all_diffs = pd.concat(diff_list, ignore_index=True).fillna(0)
                        
                        # 3. AKD ile Eşleştirme (Virman)
                        merged_list = []
                        unique_weeks = all_diffs['Week'].unique()
                        
                        for i, week in enumerate(unique_weeks):
                            if i < len(akd_files):
                                akd_info = akd_files[i]
                                df_akd = pd.read_excel(akd_info["path"])
                                
                                subset_takas = all_diffs[all_diffs['Week'] == week]
                                merged_df = df_akd.merge(subset_takas, on='Kurum', how='outer')
                                merged_list.append(merged_df)
                        
                        if merged_list:
                            final_df = pd.concat(merged_list, ignore_index=True).fillna(0)
                            final_df['Virman'] = final_df['Takas_Diff'] - final_df['Net']
                            
                            st.session_state['final_df'] = final_df
                            
                            # 4. Hacim İşleme (Opsiyonel)
                            if hacim_files:
                                hacim_df = process_hacim_files(hacim_files)
                                st.session_state['hacim_df'] = hacim_df
                                st.session_state['hacim_available'] = True
                            else:
                                st.session_state['hacim_available'] = False
                            
                            st.session_state['processed'] = True
                        else:
                            st.error("AKD dosyaları ile Takas verileri eşleştirilemedi.")
                    else:
                        st.error("Takas farkı hesaplanamadı.")

            except Exception as e:
                st.error(f"Bir hata oluştu: {str(e)}")
            
            finally:
                if takas_temp_dir and os.path.exists(takas_temp_dir):
                    shutil.rmtree(takas_temp_dir)
                if akd_temp_dir and os.path.exists(akd_temp_dir):
                    shutil.rmtree(akd_temp_dir)
                if hacim_temp_dir and os.path.exists(hacim_temp_dir):
                    shutil.rmtree(hacim_temp_dir)

# ---------------------------------------------------------
# 5. SONUÇ EKRANI
# ---------------------------------------------------------
if 'processed' in st.session_state and st.session_state['processed']:
    df = st.session_state['final_df']
    hacim_available = st.session_state.get('hacim_available', False)
    
    st.markdown("---")
    
    # Dinamik tab oluşturma
    if hacim_available:
        hacim_df = st.session_state['hacim_df']
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Virman Özet", "📋 Virman Detay", "📈 Hacim Analizi", "📉 Hacim Grafikleri"])
    else:
        tab1, tab2 = st.tabs(["📊 Virman Özet", "📋 Virman Detay"])
    
    with tab1:
        st.header("📊 Kurum Bazlı Virman Sağlaması")
        st.caption("Toplam Takas Değişimi ile Virman arasındaki farkın 0 olması beklenir.")
        
        summary_rows = []
        unique_kurumlar = sorted([str(k) for k in df['Kurum'].unique()])
        
        for kur in unique_kurumlar:
            temp = df[df['Kurum'] == kur]
            if len(temp) > 0:
                first = temp.iloc[0]['Takas_previous']
                last = temp.iloc[-1]['Takas_current']
                virman_toplam = temp['Virman'].sum()
                
                gercek_fark = last - first
                kontrol = gercek_fark - virman_toplam
                
                summary_rows.append({
                    "Kurum": kur,
                    "İlk Takas": first,
                    "Son Takas": last,
                    "Takas Değişimi": gercek_fark,
                    "Toplam Virman": virman_toplam,
                    "Fark (Kontrol)": kontrol
                })
        
        summary_df = pd.DataFrame(summary_rows)
        total_virman = summary_df['Toplam Virman'].sum()
        
        col1, col2 = st.columns([1, 3])
        with col1:
            st.metric("💰 Toplam Virman", f"{total_virman:,.0f}")
        with col2:
            min_fark = st.number_input("Sadece Farkı X'den büyük olanları göster (Mutlak)", value=0, step=100)
        
        if min_fark > 0:
            summary_df = summary_df[summary_df['Fark (Kontrol)'].abs() > min_fark]
            
        def highlight_diff(val):
            color = 'red' if abs(val) > 0 else 'green'
            return f'color: {color}; font-weight: bold'
        
        st.dataframe(
            summary_df.style.format({
                "İlk Takas": "{:,.0f}",
                "Son Takas": "{:,.0f}",
                "Takas Değişimi": "{:,.0f}",
                "Toplam Virman": "{:,.0f}",
                "Fark (Kontrol)": "{:,.2f}"
            }).applymap(highlight_diff, subset=['Fark (Kontrol)']), 
            use_container_width=True,
            height=500
        )
        
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            summary_df.to_excel(writer, index=False)
        st.download_button("📥 Virman Özet İndir", buffer.getvalue(), "Virman_Ozet.xlsx")

    with tab2:
        st.header("📋 Detaylı Virman Verileri")
        st.dataframe(df, use_container_width=True)
        
        buffer2 = io.BytesIO()
        with pd.ExcelWriter(buffer2) as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Virman Detay İndir", buffer2.getvalue(), "Virman_Detay.xlsx")

    # Hacim analizi sadece veri varsa gösterilir
    if hacim_available:
        with tab3:
            st.header("📈 Hacim Analizi")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Toplam Satır", len(hacim_df))
            with col2:
                st.metric("Kurum Sayısı", hacim_df[hacim_df['Kurum'] != 'ALL']['Kurum'].nunique())
            with col3:
                st.metric("Hafta Sayısı", hacim_df['Hafta'].nunique())
            
            st.subheader("📋 Hacim Verileri")
            st.dataframe(hacim_df, use_container_width=True)
            
            buffer3 = io.BytesIO()
            with pd.ExcelWriter(buffer3) as writer:
                hacim_df.to_excel(writer, sheet_name='Hacim', index=False)
            
            st.download_button("📥 Hacim.xlsx İndir", buffer3.getvalue(), "Hacim.xlsx")

        with tab4:
            st.header("📉 Haftalık Kurum Toplam Grafiği")
            
            plot_df = hacim_df[hacim_df['Kurum'] != 'ALL']
            
            fig = px.line(
                plot_df,
                x="Hafta",
                y="Haftalık Kurum Toplam",
                color="Kurum",
                markers=True,
                title="Haftalık Kurum Toplam Değişimi"
            )
            
            fig.update_layout(
                xaxis_title="Hafta",
                yaxis_title="Haftalık Toplam",
                legend_title="Kurum",
                hovermode="x unified",
                height=600
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.subheader("🏆 En Yüksek Hacimli 10 Kurum")
            top10 = (
                plot_df.groupby("Kurum")["Haftalık Kurum Toplam"]
                .sum()
                .sort_values(ascending=False)
                .head(10)
            )
            
            fig2 = px.bar(
                x=top10.index,
                y=top10.values,
                title="Toplam Hacim - Top 10 Kurum",
                labels={'x': 'Kurum', 'y': 'Toplam Hacim'}
            )
            
            st.plotly_chart(fig2, use_container_width=True)

else:
    st.info("""
    👋 **Hoş Geldiniz!**
    
    Bu uygulama ile Takas, Virman ve Hacim Excel dosyalarınızı analiz edebilirsiniz.
    
    **Nasıl Kullanılır:**
    1. Sol menüden ZIP dosyalarınızı yükleyin (Takas ve AKD zorunlu, Hacim opsiyonel)
    2. "Analizi Başlat" butonuna tıklayın
    3. Sonuçları görüntüleyin ve indirin
    
    **Dosya Formatları:**
    - **Takas:** Tek tarihli (örn: "1 09.xlsx", "8 09.xlsx") - ZORUNLU
    - **AKD:** Haftalık aralık (örn: "11-19 09.xlsx") - ZORUNLU
    - **Hacim:** Haftalık aralık (örn: "11-19 09.xlsx") - OPSİYONEL
    - Tüm dosyalar "Kurum" kolonu içermelidir
    - Klasör yapısı: ZIP içinde Yıl/Ay klasörleri otomatik tanınır
    """)