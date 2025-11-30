import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
import shutil
import io

# ---------------------------------------------------------
# 1. SAYFA AYARLARI
# ---------------------------------------------------------
st.set_page_config(page_title="Takas & Virman Analiz (ZIP)", page_icon="📦", layout="wide")

st.markdown("""
    <style>
    .main { padding: 2rem; }
    .stButton>button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Klasör Bazlı Takas Analizi (ZIP Yükleme)")
st.info("""
**Nasıl Kullanılır?**
1. Bilgisayarınızdaki **'takas'** klasörüne sağ tıklayıp **'ZIP dosyasına sıkıştır'** deyin.
2. Aynısını **'akd'** klasörü için yapın.
3. Oluşan ZIP dosyalarını aşağıya yükleyin. Sistem klasör yapısını (Yıl/Ay) otomatik tanıyacaktır.
""")
st.markdown("---")

# ---------------------------------------------------------
# 2. YARDIMCI FONKSİYONLAR
# ---------------------------------------------------------

def clean_takas_value(val):
    """
    Excel'den gelen veriyi sayıya çevirir.
    Örn: '1.234,56' -> 1234.56
    """
    if pd.isna(val): return 0
    if isinstance(val, (int, float)): return val
    
    # String temizliği
    val_str = str(val).strip()
    val_str = val_str.replace(".", "")  # Binlik ayracı sil
    val_str = val_str.replace(",", ".") # Ondalık ayracı nokta yap
    try:
        return float(val_str)
    except:
        return 0

def extract_zip_and_get_files(uploaded_zip, file_type="takas"):
    """
    Yüklenen ZIP dosyasını geçici bir klasöre çıkarır 
    ve içindeki dosyaları (Yıl, Ay, Gün) sırasına göre listeler.
    """
    file_list = []
    
    # Geçici klasör oluştur
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Zip dosyasını aç
        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            
        # Klasörlerde gezin (os.walk)
        for root, dirs, files in os.walk(temp_dir):
            for file in files:
                # Gereksiz dosyaları atla (Mac sistem dosyaları veya geçici excel dosyaları)
                if not file.endswith(".xlsx") or file.startswith("~$") or "__MACOSX" in root:
                    continue
                
                # Klasör yolunu parçala (Yıl ve Ay tespiti)
                rel_path = os.path.relpath(root, temp_dir)
                path_parts = rel_path.split(os.sep)
                
                # Klasör yapısını bulmaya çalış
                year = 2024 # Varsayılan
                month = 1   # Varsayılan
                
                for part in path_parts:
                    if part.isdigit():
                        val = int(part)
                        if val > 2000: # Yıl kabul et
                            year = val
                        elif 1 <= val <= 12: # Ay kabul et
                            month = val
                
                full_path = os.path.join(root, file)
                
                # Dosya isminden Gün bilgisini çekme
                try:
                    name_parts = file.replace(".xlsx", "").split()
                    
                    if file_type == "takas":
                        # Örn: "05 09.xlsx" -> ilk kısım gün
                        day = int(name_parts[0])
                        sort_key = (year, month, day)
                        display_date = f"{day}.{month}.{year}"
                        
                    elif file_type == "akd":
                        # Örn: "11-19 09.xlsx" -> ilk kısım "11-19"
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

        # Kronolojik Sıralama (Yıl -> Ay -> Gün)
        file_list.sort(key=lambda x: x["sort_key"])
        return file_list, temp_dir

    except zipfile.BadZipFile:
        st.error("Yüklenen dosya geçerli bir ZIP dosyası değil.")
        return [], None
    except Exception as e:
        st.error(f"ZIP açılırken hata oluştu: {e}")
        return [], None

# ---------------------------------------------------------
# 3. SIDEBAR (DOSYA YÜKLEME)
# ---------------------------------------------------------
with st.sidebar:
    st.header("📂 Dosya Yükleme")
    
    st.subheader("1️⃣ Takas Klasörü (ZIP)")
    takas_zip = st.file_uploader("Takas.zip yükleyin", type="zip")
    
    st.subheader("2️⃣ AKD Klasörü (ZIP)")
    akd_zip = st.file_uploader("AKD.zip yükleyin", type="zip")
    
    st.markdown("---")
    process_button = st.button("🚀 Analizi Başlat", type="primary")

# ---------------------------------------------------------
# 4. İŞLEM MANTIĞI
# ---------------------------------------------------------
if process_button:
    if not takas_zip or not akd_zip:
        st.error("❌ Lütfen her iki ZIP dosyasını da yükleyin.")
    else:
        # Geçici değişkenler
        takas_temp_dir = None
        akd_temp_dir = None
        
        with st.spinner("📦 ZIP dosyaları açılıyor ve analiz ediliyor..."):
            try:
                # 1. Dosyaları Çıkar
                takas_files, takas_temp_dir = extract_zip_and_get_files(takas_zip, "takas")
                akd_files, akd_temp_dir = extract_zip_and_get_files(akd_zip, "akd")
                
                if not takas_files or not akd_files:
                    st.error("ZIP içeriğinde uygun Excel dosyaları bulunamadı. Lütfen klasör yapısını kontrol edin.")
                else:
                    st.success(f"✅ {len(takas_files)} Takas ve {len(akd_files)} AKD dosyası bulundu.")
                    
                    # 2. Takas Farklarını Hesapla
                    diff_list = []
                    # i=1'den başlıyoruz çünkü bir önceki dosyayla kıyaslayacağız
                    for i in range(1, len(takas_files)):
                        prev = takas_files[i - 1]
                        curr = takas_files[i]
                        
                        df_prev = pd.read_excel(prev["path"])
                        df_curr = pd.read_excel(curr["path"])
                        
                        # Temizlik
                        df_prev["Takas"] = df_prev["Takas"].apply(clean_takas_value)
                        df_curr["Takas"] = df_curr["Takas"].apply(clean_takas_value)
                        
                        # Birleştir
                        df_merged = pd.merge(
                            df_curr, df_prev, on="Kurum", 
                            suffixes=("_current", "_previous"), how="outer"
                        )
                        
                        # Fark Hesabı
                        df_merged["Takas_Diff"] = df_merged["Takas_current"].fillna(0) - df_merged["Takas_previous"].fillna(0)
                        
                        # Hafta Etiketi
                        df_merged["Week"] = f"{prev['display']} - {curr['display']}"
                        
                        diff_list.append(df_merged)
                    
                    if diff_list:
                        all_diffs = pd.concat(diff_list, ignore_index=True).fillna(0)
                        
                        # 3. AKD ile Eşleştirme (Virman Hesabı)
                        merged_list = []
                        unique_weeks = all_diffs['Week'].unique()
                        
                        # Takas haftaları ile AKD dosyalarını sırasıyla eşleştir
                        for i, week in enumerate(unique_weeks):
                            if i < len(akd_files):
                                akd_info = akd_files[i]
                                df_akd = pd.read_excel(akd_info["path"])
                                
                                # İlgili haftanın takas farkları
                                subset_takas = all_diffs[all_diffs['Week'] == week]
                                
                                # Takas ve AKD birleştir
                                merged_df = df_akd.merge(subset_takas, on='Kurum', how='outer')
                                merged_list.append(merged_df)
                        
                        if merged_list:
                            final_df = pd.concat(merged_list, ignore_index=True).fillna(0)
                            
                            # Virman Formülü: (Takas Farkı - Net Alım)
                            final_df['Virman'] = final_df['Takas_Diff'] - final_df['Net']
                            
                            st.session_state['final_df'] = final_df
                            st.session_state['processed'] = True
                        else:
                            st.error("AKD dosyaları ile Takas verileri eşleştirilemedi.")
                    else:
                        st.error("Takas farkı hesaplanamadı (En az 2 sıralı dosya gerekir).")

            except Exception as e:
                st.error(f"Bir hata oluştu: {str(e)}")
            
            finally:
                # 4. Temizlik (Geçici klasörleri sil)
                if takas_temp_dir and os.path.exists(takas_temp_dir):
                    shutil.rmtree(takas_temp_dir)
                if akd_temp_dir and os.path.exists(akd_temp_dir):
                    shutil.rmtree(akd_temp_dir)

# ---------------------------------------------------------
# 5. SONUÇ EKRANI
# ---------------------------------------------------------
if 'processed' in st.session_state and st.session_state['processed']:
    df = st.session_state['final_df']
    
    st.markdown("### 📊 Analiz Sonuçları")
    
    tab1, tab2 = st.tabs(["Özet & Virman Kontrol", "Tüm Veriler"])
    
    with tab1:
        st.write("**Kurum Bazlı Virman Sağlaması**")
        st.caption("Toplam Takas Değişimi ile Virman arasındaki farkın 0 olması beklenir.")
        
        # Özet Tabloyu Hazırla
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
        
        # Toplam Virman Gösterimi
        total_virman = summary_df['Toplam Virman'].sum()
        st.markdown(f"### 💰 Toplam Virman: {total_virman:,.0f}")
        
        # Filtreleme
        col1, col2 = st.columns(2)
        with col1:
            min_fark = st.number_input("Sadece Farkı X'den büyük olanları göster (Mutlak)", value=0, step=100)
        
        if min_fark > 0:
            summary_df = summary_df[summary_df['Fark (Kontrol)'].abs() > min_fark]
            
        # TABLO GÖSTERİMİ (Renkli Fark Kontrol)
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
        
        # Excel İndir
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer) as writer:
            summary_df.to_excel(writer, index=False)
        st.download_button("📥 Özet Raporu İndir", buffer.getvalue(), "Virman_Ozet.xlsx")


    with tab2:
        st.write("Tüm haftaların birleştirilmiş detaylı verisi:")
        st.dataframe(df, use_container_width=True)
        
        buffer2 = io.BytesIO()
        with pd.ExcelWriter(buffer2) as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Detaylı Veriyi İndir", buffer2.getvalue(), "Virman_Detay.xlsx")