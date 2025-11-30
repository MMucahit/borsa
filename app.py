import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime

st.set_page_config(page_title="Takas & Hacim Analiz", page_icon="📊", layout="wide")

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.title("📊 Takas & Hacim Analiz Sistemi")
st.markdown("---")

# Sidebar
with st.sidebar:
    st.header("📁 Dosya Yükleme")
    st.markdown("Excel dosyalarınızı yükleyin")
    
    st.subheader("1️⃣ Takas Dosyaları")
    st.caption("Format: 1 09.xlsx, 8 09.xlsx")
    takas_files = st.file_uploader(
        "Takas", 
        type=['xlsx'], 
        accept_multiple_files=True,
        key="takas"
    )
    
    st.subheader("2️⃣ AKD Dosyaları")
    st.caption("Format: 11-19 09.xlsx")
    akd_files = st.file_uploader(
        "AKD", 
        type=['xlsx'], 
        accept_multiple_files=True,
        key="akd"
    )
    
    st.subheader("3️⃣ Hacim Dosyaları")
    st.caption("Format: 11-19 09.xlsx")
    hacim_files = st.file_uploader(
        "Hacim", 
        type=['xlsx'], 
        accept_multiple_files=True,
        key="hacim"
    )
    
    st.markdown("---")
    process_button = st.button("🚀 Analizi Başlat", type="primary")

# Helper Functions
def parse_date(filename):
    """Parse filename like '1 09.xlsx' -> (month, day)"""
    name = filename.replace('.xlsx', '')
    day, month = name.split()
    return int(month), int(day)

def parse_week_filename(filename):
    """Parse filename like '11-19 09.xlsx' -> (month, start_day)"""
    name = filename.replace('.xlsx', '')
    day_range, month = name.split()
    start_day, _ = day_range.split('-')
    return int(month), int(start_day)

def process_takas_files(takas_files):
    """Process Takas files and calculate differences"""
    # Sort files
    files_dict = {f.name: f for f in takas_files}
    sorted_names = sorted(files_dict.keys(), key=parse_date)
    
    diff_list = []
    
    for i in range(1, len(sorted_names)):
        prev_name = sorted_names[i - 1]
        curr_name = sorted_names[i]
        
        df_prev = pd.read_excel(files_dict[prev_name])
        df_curr = pd.read_excel(files_dict[curr_name])
        
        # Sadece numeric olması gereken kolonları dönüştür
        df_prev["Takas"] = (
            df_prev["Takas"]
            .astype(str)
            .str.replace(".", "", regex=False)   # binlik ayırıcı noktaları sil
            .str.replace(",", ".", regex=False)  # ondalık virgülü noktaya çevir
            .pipe(pd.to_numeric, errors="coerce")
        )

        df_curr["Takas"] = (
            df_curr["Takas"]
            .astype(str)
            .str.replace(".", "", regex=False)
            .str.replace(",", ".", regex=False)
            .pipe(pd.to_numeric, errors="coerce")
        )

        # Merge by Kurum
        df_merged = pd.merge(
            df_curr,
            df_prev,
            on="Kurum",
            suffixes=("_current", "_previous"),
            how="outer"
        )
        
        # Calculate difference
        df_merged["Takas_Diff"] = (
            df_merged["Takas_current"].fillna(0) - 
            df_merged["Takas_previous"].fillna(0)
        )
        
        # Label the week
        df_merged["Week"] = f"{prev_name.replace('.xlsx', '')} - {curr_name.replace('.xlsx', '')}"
        
        diff_list.append(df_merged)
    
    # Combine all differences
    all_diffs = pd.concat(diff_list, ignore_index=True).fillna(0)
    
    # Ensure Week column is string type
    all_diffs['Week'] = all_diffs['Week'].astype(str)
    
    return all_diffs, sorted_names

def process_virman(all_diffs, akd_files, sorted_takas_names):
    """Merge Takas differences with AKD data and calculate Virman"""
    # Sort AKD files
    akd_dict = {f.name: f for f in akd_files}
    sorted_akd = sorted(akd_dict.keys(), key=parse_week_filename)
    
    merged_list = []
    
    for i, week in enumerate(all_diffs['Week'].unique()):
        if i < len(sorted_akd):
            df_akd = pd.read_excel(akd_dict[sorted_akd[i]])
            merged_df = df_akd.merge(
                all_diffs[all_diffs['Week'] == week],
                on='Kurum',
                how='outer'
            )
            merged_list.append(merged_df)
    
    final_df = pd.concat(merged_list, ignore_index=True).fillna(0)
    
    # Calculate Virman
    final_df['Virman'] = final_df['Takas_Diff'] - final_df['Net']
    
    # Ensure all columns are proper types
    final_df['Week'] = final_df['Week'].astype(str)
    final_df['Kurum'] = final_df['Kurum'].astype(str)
    
    return final_df

def process_hacim_files(hacim_files):
    """Process Hacim files and calculate percentages"""
    # Sort files
    hacim_dict = {f.name: f for f in hacim_files}
    sorted_names = sorted(hacim_dict.keys(), key=parse_week_filename)
    
    all_data = []
    
    for file_name in sorted_names:
        df = pd.read_excel(hacim_dict[file_name])
        
        # Group by Kurum and sum
        grouped = df.groupby("Kurum", as_index=False)["Toplam"].sum()
        
        # Calculate weekly total
        grand_total = grouped["Toplam"].sum()
        
        # Add percentage
        grouped["Yüzde"] = (grouped["Toplam"] / grand_total * 100).round(2)
        
        # Rename
        grouped.rename(columns={"Toplam": "Haftalık Kurum Toplam"}, inplace=True)
        
        # Add weekly total
        grouped["Haftalık Toplam"] = grand_total
        
        # Add week name
        grouped["Hafta"] = file_name.replace('.xlsx', '')
        
        # Add ALL row
        all_row = pd.DataFrame({
            "Kurum": ["ALL"],
            "Haftalık Kurum Toplam": [grand_total],
            "Yüzde": [100.0],
            "Haftalık Toplam": [grand_total],
            "Hafta": [file_name.replace('.xlsx', '')]
        })
        
        grouped = pd.concat([grouped, all_row], ignore_index=True)
        all_data.append(grouped)
    
    final_df = pd.concat(all_data, ignore_index=True)
    return final_df

# Main Processing Logic
if process_button:
    if not takas_files or not akd_files or not hacim_files:
        st.error("❌ Lütfen tüm dosyaları yükleyin!")
    else:
        with st.spinner("⏳ Dosyalar işleniyor..."):
            try:
                # Process Takas
                st.info("📊 Takas dosyaları işleniyor...")
                all_diffs, sorted_takas = process_takas_files(takas_files)
                
                # Process Virman
                st.info("💰 Virman hesaplanıyor...")
                virman_df = process_virman(all_diffs, akd_files, sorted_takas)
                
                # Process Hacim
                st.info("📈 Hacim analizi yapılıyor...")
                hacim_df = process_hacim_files(hacim_files)
                
                # Store in session state
                st.session_state['virman_df'] = virman_df
                st.session_state['hacim_df'] = hacim_df
                st.session_state['processed'] = True
                
                st.success("✅ İşlem başarıyla tamamlandı!")
                
            except Exception as e:
                st.error(f"❌ Hata oluştu: {str(e)}")

# Display Results
if 'processed' in st.session_state and st.session_state['processed']:
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["📊 Virman Analizi", "📈 Hacim Analizi", "📉 Grafikler"])
    
    with tab1:
        st.header("💰 Virman Analizi")
        
        virman_df = st.session_state['virman_df']
        
        # Toplam Virman hesapla
        total_virman = virman_df['Virman'].sum()
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Toplam Satır", len(virman_df))
        with col2:
            st.metric("Kurum Sayısı", virman_df['Kurum'].nunique())
        with col3:
            st.metric("Hafta Sayısı", virman_df['Week'].nunique())
        with col4:
            st.metric("💰 TOPLAM VİRMAN", f"{total_virman:,.2f}")
        
        st.markdown("---")
        
        # Filtreler
        st.subheader("🔍 Filtreler")
        col_filter1, col_filter2 = st.columns(2)
        
        with col_filter1:
            kurum_options = sorted([str(k) for k in virman_df['Kurum'].unique()])
            selected_kurum = st.multiselect(
                "Kurum Seçin",
                options=kurum_options,
                default=None,
                placeholder="Tüm Kurumlar"
            )
        
        with col_filter2:
            week_options = sorted([str(w) for w in virman_df['Week'].unique()])
            selected_week = st.multiselect(
                "Hafta Seçin",
                options=week_options,
                default=None,
                placeholder="Tüm Haftalar"
            )
        
        # Filtreleme uygula
        filtered_df = virman_df.copy()
        
        if selected_kurum:
            filtered_df = filtered_df[filtered_df['Kurum'].isin(selected_kurum)]
        
        if selected_week:
            filtered_df = filtered_df[filtered_df['Week'].isin(selected_week)]
        
        # Filtrelenmiş toplam
        filtered_virman = filtered_df['Virman'].sum()
        
        if selected_kurum or selected_week:
            st.info(f"📊 Filtrelenmiş Toplam Virman: **{filtered_virman:,.2f}**")
        
        st.subheader("📋 Virman Verileri")
        st.dataframe(filtered_df, width="stretch")
        
        # Kurum bazlı özet
        st.subheader("🏢 Kurum Bazlı Özet")
        summary_list = []
        for kurum in virman_df['Kurum'].unique():
            kurum_data = virman_df[virman_df['Kurum'] == kurum]
            if len(kurum_data) > 0:
                first_takas = kurum_data.iloc[0]['Takas_previous']
                last_takas = kurum_data.iloc[-1]['Takas_current']
                total_virman = kurum_data['Virman'].sum()
                fark = (last_takas - first_takas) - total_virman
                
                summary_list.append({
                    'Kurum': kurum,
                    'İlk Takas': first_takas,
                    'Son Takas': last_takas,
                    'Toplam Virman': total_virman,
                    'Fark': fark
                })
        
        summary_df = pd.DataFrame(summary_list)
        st.dataframe(summary_df, width="stretch")
        
        # Download button
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            virman_df.to_excel(writer, sheet_name='Virman', index=False)
        
        st.download_button(
            label="📥 Virman.xlsx İndir",
            data=buffer.getvalue(),
            file_name="Virman.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with tab2:
        st.header("📈 Hacim Analizi")
        
        hacim_df = st.session_state['hacim_df']
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Toplam Satır", len(hacim_df))
        with col2:
            st.metric("Kurum Sayısı", hacim_df[hacim_df['Kurum'] != 'ALL']['Kurum'].nunique())
        with col3:
            st.metric("Hafta Sayısı", hacim_df['Hafta'].nunique())
        
        st.subheader("📋 Hacim Verileri")
        st.dataframe(hacim_df, width="stretch")
        
        # Download button
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            hacim_df.to_excel(writer, sheet_name='Hacim', index=False)
        
        st.download_button(
            label="📥 Hacim.xlsx İndir",
            data=buffer.getvalue(),
            file_name="Hacim.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with tab3:
        st.header("📉 Haftalık Kurum Toplam Grafiği")
        
        hacim_df = st.session_state['hacim_df']
        
        # Filter out ALL
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
        
        # Top 10 Institutions
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
    # Welcome message
    st.info("""
    👋 **Hoş Geldiniz!**
    
    Bu uygulama ile Takas ve Hacim Excel dosyalarınızı analiz edebilirsiniz.
    
    **Nasıl Kullanılır:**
    1. Sol menüden dosyalarınızı yükleyin
    2. "Analizi Başlat" butonuna tıklayın
    3. Sonuçları görüntüleyin ve indirin
    
    **Dosya Formatları:**
    - **Takas:** Tek tarihli (örn: "1 09.xlsx", "8 09.xlsx")
    - **AKD & Hacim:** Haftalık aralık (örn: "11-19 09.xlsx")
    - Tüm dosyalar "Kurum" kolonu içermelidir
    """)