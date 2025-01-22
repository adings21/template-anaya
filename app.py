import streamlit as st
import pandas as pd
from fpdf import FPDF
import os
from datetime import datetime
from num2words import num2words

# Fungsi untuk membuat laporan PDF
class PDF(FPDF):
    def header(self):
        self.image('Logo-Anaya.jpg', 10, 8, 33)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'PT. ANAYA GLOBAL INDONESIA', 0, 1, 'C')
        self.set_font('Arial', '', 10)
        self.cell(0, 6, 'Ruko Paramount 7 - CS - DF -2/18 RT 02 RW 02 Curug', 0, 1, 'C')
        self.cell(0, 6, 'Sangereng - Kelapa Dua Tangerang', 0, 1, 'C')
        self.cell(0, 6, '02129324688 | accounting@anaya.co.id', 0, 1, 'C')
        self.cell(0, 6, 'www.anaya.co.id | 3142998258451000', 0, 1, 'C')
        self.ln(10)

    def footer(self):
        self.set_y(-30)
        self.set_font('Arial', '', 10)
        self.cell(0, 6, "Iswandi Simardjo", 0, 1, 'R')
        self.cell(0, 6, "PT. Anaya Global Indonesia", 0, 1, 'R')
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def generate_pdf(data, output_file):
    pdf = PDF()
    pdf.add_page()

    pdf.set_font('Arial', 'B', 9)
    pdf.cell(130)
    pdf.cell(0, 6, f"Tanggal: {data['Tanggal']}", 0, 1 ,'R')
    pdf.cell(130)
    pdf.cell(0, 6, f"Faktur #: {data['Nomor Transaksi']}", 0, 1,'R')
    pdf.cell(130)
    pdf.cell(0, 6, f"Referensi Pelanggan: {data['No Ref']}", 0, 1, 'R')
    pdf.ln(10)

    pdf.set_font('Arial', 'B', 10)
    pdf.cell(0, 6, "PELANGGAN:", 0, 1, 'L')
    pdf.set_font('Arial', '', 9)
    pdf.multi_cell(0, 6, data['Nama Perusahaan'], 0, 'L')
    pdf.ln(2)
    pdf.multi_cell(0, 6, data['Alamat Penagihan'], 0, 'L')
    pdf.ln(2)
    pdf.multi_cell(0, 6, data['Alamat Pengiriman'], 0, 'L')
    pdf.ln(10)

    # Header Tabel
    pdf.set_fill_color(200, 200, 200)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, "Tag", 1, 0, 'C', True)
    pdf.cell(50, 6, "Metode Pembayaran", 1, 0, 'C', True)
    pdf.cell(40, 6, "Jatuh Tempo", 1, 0, 'C', True)
    pdf.cell(40, 6, "Tanggal Pengiriman", 1, 1, 'C', True)

    pdf.set_font('Arial', '', 8)
    pdf.cell(30, 6, data["Tag"][:20], 1, 0, 'C')
    pdf.cell(50, 6, data["Metode Pembayaran"][:30], 1, 0, 'C')
    pdf.cell(40, 6, data["Tanggal Jatuh Tempo"], 1, 0, 'C')
    pdf.cell(40, 6, data["Tanggal"], 1, 1, 'C')
    pdf.ln(5)

        # Detail Tabel
    pdf.set_fill_color(200, 200, 200)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(15, 6, 'No.', 1, 0, 'C', True)
    pdf.cell(20, 6, 'Qty', 1, 0, 'C', True)
    pdf.cell(50, 6, 'Nama Barang', 1, 0, 'C', True)
    pdf.cell(30, 6, 'Harga Satuan', 1, 0, 'C', True)
    pdf.cell(25, 6, 'Diskon', 1, 0, 'C', True)
    pdf.cell(25, 6, 'Pajak', 1, 0, 'C', True)
    pdf.cell(30, 6, 'Jumlah', 1, 1, 'C', True)

    # Isi Tabel
    pdf.set_font('Arial', '', 8)
    pdf.set_fill_color(255, 255, 255)

    # Inisialisasi nomor urut
    no = 1  
    for _, row in data['detail'].iterrows():
        pdf.cell(15, 6, str(no), 1)  # Kolom No.
        pdf.cell(20, 6, str(row['Kuantitas']), 1)  # Kolom Qty
        pdf.cell(50, 6, row['Deskripsi'][:30], 1)  # Nama Barang (max 30 karakter)
        pdf.cell(30, 6, f"{row['Harga per Unit']:,}", 1)  # Harga Satuan
        pdf.cell(25, 6, f"{row['Diskon Per Baris %']}", 1)  # Diskon
        pdf.cell(25, 6, f"PPN{row['Tarif Pajak']}%", 1)  # Pajak
        pdf.cell(30, 6, f"{row['Jumlah Kena Pajak per Baris']:,}", 1, 1)  # Jumlah
        no += 1  # Tambahkan nomor urut
    pdf.ln(5)


   # Subtotal dan lainnya
    pdf.set_font('Arial', 'B', 9)
    pdf.set_fill_color(240, 240, 240)  # Warna abu terang
    # Subtotal
    pdf.cell(120)  # Geser ke kanan
    pdf.cell(30, 10, 'Subtotal:', 0, 0, 'L', True)  # Warna dengan fill
    pdf.cell(40, 10, f"{data['subtotal']:,}", 0, 1, 'R', True)
    # DPP (11/12)
    dpp = data['subtotal'] * 11 / 12  # Perhitungan DPP
    pdf.set_fill_color(255, 255, 255)  # Warna putih (tidak ada fill)
    pdf.cell(120)
    pdf.cell(30, 10, 'DPP (11/12):', 0, 0, 'L', True)
    pdf.cell(40, 10, f"{dpp:,.2f}", 0, 1, 'R', True)
    # PPN 12%
    ppn_12 = dpp * 0.12  # Perhitungan PPN 12%
    pdf.set_fill_color(240, 240, 240)  # Warna abu terang
    pdf.cell(120)
    pdf.cell(30, 10, 'PPN 12%:', 0, 0, 'L', True)
    pdf.cell(40, 10, f"{ppn_12:,.2f}", 0, 1, 'R', True)
    # Total
    total = data['subtotal'] + ppn_12
    pdf.set_fill_color(255, 255, 255)  # Warna putih (tidak ada fill)
    pdf.cell(120)
    pdf.cell(30, 10, 'Total:', 0, 0, 'L', True)  # Warna dengan fill
    pdf.cell(40, 10, f"{total:,.2f}", 0, 1, 'R', True)
    # Sisa Tagihan
    pdf.set_fill_color(240, 240, 240)  # Warna abu terang
    pdf.cell(120)
    pdf.cell(30, 10, 'Sisa Tagihan:', 0, 0, 'L', True)  # Warna tanpa fill
    pdf.cell(40, 10, f"{data['Sisa Tagihan']:,}", 0, 1, 'R', True)


    # TERBILANG dan Detail Pembayaran
    terbilang_text = f"TERBILANG: {data['terbilang']}"
    # Posisi "TERBILANG" ke kiri, sejajar dengan Subtotal
    pdf.set_xy(10, pdf.get_y() + 5)  # Posisikan sedikit di bawah subtotal
    pdf.set_font('Arial', 'B', 9)  # Font Bold untuk TERBILANG
    pdf.multi_cell(100, 6, terbilang_text, 0, 'L')  # MultiCell untuk teks panjang
    pdf.ln(2)  # Tambahkan jarak vertikal
    # Detail Pembayaran
    pdf.set_font('Arial', 'B', 9)  # Font Bold untuk judul Detail Pembayaran
    pdf.cell(0, 6, "Detail Pembayaran:", 0, 1, 'L')
    pdf.set_font('Arial', '', 8)  # Font biasa untuk detail pembayaran
    pdf.cell(0, 6, "Nama Bank : BCA", 0, 1, 'L')
    pdf.cell(0, 6, "Cabang Bank : KCP MANDALA RAYA", 0, 1, 'L')
    pdf.cell(0, 6, "Nomor Akun Bank : 3983101018", 0, 1, 'L')
    pdf.cell(0, 6, "Atas Nama : PT ANAYA GLOBAL INDONESIA", 0, 1, 'L')

    pdf.output(output_file)

# Streamlit App
st.title("Laporan PDF Generator - PT. Anaya Global Indonesia")

uploaded_file = st.file_uploader("Unggah file Excel Anda", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    transaction_ids = df['Nomor Transaksi'].unique()

    selected_transaction = st.selectbox("Pilih Nomor Transaksi", transaction_ids)

    metode_pembayaran = st.text_input("Metode Pembayaran")
    if st.button("Generate PDF"):
        selected_data = df[df['Nomor Transaksi'] == selected_transaction]
        subtotal = selected_data['Jumlah Kena Pajak per Baris'].sum()
        ppn = subtotal * 0.11
        total = subtotal + ppn
        sisa_tagihan = total
        terbilang = num2words(sisa_tagihan, lang='id').capitalize()

        report_data = {
            "Tanggal": selected_data.iloc[0]['Tanggal'],
            "Nomor Transaksi": selected_transaction,
            "No Ref": selected_data.iloc[0]['No Ref'],
            "Nama Perusahaan": selected_data.iloc[0]['Nama Perusahaan'],
            "Alamat Penagihan": selected_data.iloc[0]['Alamat Penagihan'],
            "Alamat Pengiriman": selected_data.iloc[0]['Alamat Pengiriman'],
            "Tag": selected_data.iloc[0]['Tag'],
            "Metode Pembayaran": metode_pembayaran,
            "Tanggal Jatuh Tempo": selected_data.iloc[0]['Tanggal Jatuh Tempo'],
            "Tanggal Pengiriman": selected_data.iloc[0]['Tanggal'],
            "detail": selected_data,
            "subtotal": subtotal,
            "ppn": ppn,
            "total": total,
            "Sisa Tagihan": sisa_tagihan,
            "terbilang": terbilang,
        }

        output_file = f"Laporan_template_anaya.pdf"
        generate_pdf(report_data, output_file)
        st.success(f"Laporan berhasil dibuat: {output_file}")
        with open(output_file, "rb") as file:
            st.download_button(label="Unduh PDF", data=file, file_name=output_file, mime="application/pdf")
        os.remove(output_file)
