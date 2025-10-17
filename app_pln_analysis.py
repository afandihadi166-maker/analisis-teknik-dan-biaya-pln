import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(page_title="Simulasi RAB & Efisiensi PLN", layout="wide", page_icon="⚡")

# -------------------------------
# HEADER
# -------------------------------
st.markdown(
    """
    <h1 style='text-align:center; color:#0047AB;'>⚡ Simulasi Efisiensi & Analisis RAB PLN ULP Batang ⚡</h1>
    <p style='text-align:center;'>Program ini membaca file <b>RAB PLN (2 Sheet: RAB & Gambar)</b> untuk menghitung:</p>
    <ul>
        <li>Total biaya per komponen</li>
        <li>Rugi daya (losses I²R) berdasarkan jenis kabel dan panjang</li>
        <li>Efisiensi teknis & trafo</li>
        <li>ROI proyek (Return on Investment)</li>
    </ul>
    """,
    unsafe_allow_html=True
)

# -------------------------------
# SIDEBAR
# -------------------------------
st.sidebar.header("⚙️ Pengaturan Asumsi")
tarif_kwh = st.sidebar.number_input("Tarif Listrik (Rp/kWh)", value=1500, min_value=0)
faktor_daya = st.sidebar.number_input("Faktor Daya", value=0.8, min_value=0.0, max_value=1.0)
tipe_phase_default = st.sidebar.selectbox("Asumsi Default Tipe Phase", ["3 Phase", "1 Phase"], index=0)

resistansi_kabel = {
    "NYY 3x70 mm²": 0.268,
    "NYY 3x35 mm²": 0.524,
    "NFA2X-T 2x70 + N70 mm²": 0.4,
}

# -------------------------------
# UPLOAD FILE
# -------------------------------
uploaded_file = st.file_uploader("📂 Unggah File Excel Template RAB PLN (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        required_sheets = ["RAB", "Gambar"]
        if not all(sheet in excel_file.sheet_names for sheet in required_sheets):
            st.error(f"Sheet yang diperlukan: {required_sheets}.")
            st.stop()

        df_rab = pd.read_excel(uploaded_file, sheet_name="RAB")
        df_gambar = pd.read_excel(uploaded_file, sheet_name="Gambar")

        required_columns_rab = ["Total (Rp)"]
        required_columns_gambar = ["Nama Lokasi", "Jenis Kabel", "Panjang Jaringan (m)", "Beban Total (kVA)", "Tegangan (V)"]

        if not all(col in df_rab.columns for col in required_columns_rab):
            st.error("Kolom wajib di sheet 'RAB' tidak lengkap.")
            st.stop()

        if not all(col in df_gambar.columns for col in required_columns_gambar):
            st.error("Kolom wajib di sheet 'Gambar' tidak lengkap.")
            st.stop()

        if "Tipe Phase" in df_gambar.columns:
            df_gambar["Tipe Phase"] = df_gambar["Tipe Phase"].str.upper().replace({"1 PHASE": "1 Phase", "3 PHASE": "3 Phase"})
        else:
            st.warning(f"Kolom 'Tipe Phase' tidak ada, diasumsikan '{tipe_phase_default}'.")
            df_gambar["Tipe Phase"] = tipe_phase_default

        # Bersihkan data
        df_gambar = df_gambar.fillna({
            "Jenis Kabel": "-",
            "Panjang Jaringan (m)": 0,
            "Beban Total (kVA)": 0,
            "Tegangan (V)": 380,
        })

        st.subheader("📘 Data RAB")
        st.dataframe(df_rab, use_container_width=True)

        st.subheader("📐 Data Teknis Jaringan")
        st.dataframe(df_gambar, use_container_width=True)

        # -------------------------------
        # PERHITUNGAN
        # -------------------------------
        def hitung_losses(baris):
            jenis = baris.get("Jenis Kabel", "-")
            panjang = baris.get("Panjang Jaringan (m)", 0)
            beban = baris.get("Beban Total (kVA)", 0)
            tegangan = baris.get("Tegangan (V)", 380)
            tipe_phase = baris.get("Tipe Phase", "3 Phase").upper()

            if jenis not in resistansi_kabel or beban <= 0:
                return 0

            r = resistansi_kabel[jenis] * (panjang / 1000)

            if tipe_phase == "3 PHASE":
                i = (beban * 1000) / (np.sqrt(3) * tegangan)
            else:
                i = (beban * 1000) / tegangan

            p_loss = (i ** 2) * r
            return p_loss / 1000  # kW

        df_gambar["Losses (kW)"] = df_gambar.apply(hitung_losses, axis=1)
        df_gambar["Efisiensi (%)"] = 100 * (1 - (df_gambar["Losses (kW)"] / ((df_gambar["Beban Total (kVA)"] * faktor_daya) + 1e-6)))
        df_gambar["Manfaat (Rp/tahun)"] = df_gambar["Losses (kW)"] * 8760 * tarif_kwh

        total_biaya = df_rab["Total (Rp)"].sum()
        total_manfaat = df_gambar["Manfaat (Rp/tahun)"].sum()
        roi = (total_manfaat / total_biaya) * 100 if total_biaya > 0 else 0

        # -------------------------------
        # HASIL
        # -------------------------------
        st.subheader("📊 Hasil Analisis Teknis & Ekonomi")
        lokasi_options = ['Semua'] + list(df_gambar['Nama Lokasi'].unique())
        selected_lokasi = st.selectbox("Pilih Lokasi untuk Analisis", lokasi_options)

        df_filtered = df_gambar if selected_lokasi == 'Semua' else df_gambar[df_gambar['Nama Lokasi'] == selected_lokasi]

        st.dataframe(
            df_filtered[["Nama Lokasi", "Tipe Phase", "Losses (kW)", "Efisiensi (%)", "Manfaat (Rp/tahun)"]]
            .style.format({
                "Losses (kW)": "{:.2f}",
                "Efisiensi (%)": "{:.2f}",
                "Manfaat (Rp/tahun)": "Rp {:,.0f}"
            }),
            use_container_width=True
        )

        # Grafik interaktif pakai Plotly
        st.subheader("📉 Visualisasi Losses per Lokasi")
        fig = px.bar(
            df_filtered,
            x="Nama Lokasi",
            y="Losses (kW)",
            color="Losses (kW)",
            color_continuous_scale="blues",
            title="Grafik Rugi Daya (kW)",
            height=450
        )
        st.plotly_chart(fig, use_container_width=True)

        col1, col2 = st.columns(2)
        col1.metric("💰 Total Biaya RAB", f"Rp {total_biaya:,.0f}")
        col2.metric("📈 ROI Tahunan", f"{roi:.2f}%")

        csv = df_filtered.to_csv(index=False)
        st.download_button(
            label="📥 Unduh Hasil Analisis (CSV)",
            data=csv,
            file_name="hasil_analisis_rab_pln.csv",
            mime="text/csv",
        )

        # -------------------------------
        # KESIMPULAN BERWARNA
        # -------------------------------
        avg_efisiensi = df_filtered["Efisiensi (%)"].mean()
        avg_losses = df_filtered["Losses (kW)"].mean()

        st.subheader(" Kesimpulan")

        kesimpulan_blocks = []

        # Efisiensi
        if avg_efisiensi > 95:
            color1 = "#65e23b"  # hijau muda
            teks1 = f"Efisiensi sistem sangat baik ({avg_efisiensi:.2f}%)."
        elif avg_efisiensi > 90:
            color1 = "#e6bf42"  # kuning muda
            teks1 = f"Efisiensi sistem baik ({avg_efisiensi:.2f}%)."
        else:
            color1 = "#f8d7da"  # merah muda
            teks1 = f"Efisiensi sistem perlu ditingkatkan ({avg_efisiensi:.2f}%)."

        # Losses
        if avg_losses < 1:
            color2 = "#68e240"
            teks2 = f"Rugi daya sangat rendah ({avg_losses:.2f} kW), jaringan efisien."
        elif avg_losses < 3:
            color2 = "#d8b43f"
            teks2 = f"Rugi daya {avg_losses:.2f} kW, masih dalam batas wajar."
        else:
            color2 = "#f8d7da"
            teks2 = f"Rugi daya tinggi ({avg_losses:.2f} kW), perlu evaluasi kabel dan beban."

        # ROI
        if roi > 20:
            color3 = "#77f04f"
            teks3 = f"Proyek ini layak secara ekonomi (ROI {roi:.2f}%)."
        elif roi > 10:
            color3 = "#ecc852"
            teks3 = f"Proyek cukup layak (ROI {roi:.2f}%)."
        else:
            color3 = "#f12738"
            teks3 = f"Proyek kurang layak (ROI {roi:.2f}%)."

        # Tampilkan blok warna
        st.markdown(f"<div style='background-color:{color1}; padding:10px; border-radius:10px; margin-bottom:8px;'>{teks1}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='background-color:{color2}; padding:10px; border-radius:10px; margin-bottom:8px;'>{teks2}</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='background-color:{color3}; padding:10px; border-radius:10px;'>{teks3}</div>", unsafe_allow_html=True)

        st.success("✅ Analisis selesai! Semua perhitungan dan kesimpulan telah ditampilkan.")

        st.markdown("---")
        st.caption("Dibuat dengan 💡 Streamlit | PLN ULP Batang © 2025")

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")

else:
    st.info("📥 Silakan unggah file Excel RAB PLN terlebih dahulu untuk memulai analisis.")
