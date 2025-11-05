import pandas as pd
import os
import numpy as np
import matplotlib.pyplot as plt

# ===============================
# 1️⃣ SETUP & IMPORT DATA
# ===============================
os.chdir("C:/Users/Lab. TI_GKB/Desktop")
file_path = "Data_Wisudawan.xlsx"
output_file = "rekap_wisuda_final.xlsx"

try:
    # Periksa apakah file ada
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File '{file_path}' tidak ditemukan!")

    # Baca file Excel
    data = pd.read_excel(file_path)

    # ===============================
    # 2️⃣ DATA CLEANSING
    # ===============================
    data.columns = data.columns.str.strip()
    data = data.dropna(subset=['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Lama Studi (Semester)', 'Tahun Wisuda'])
    data = data.fillna({'IPK': 0, 'Lama Studi (Semester)': 0})

    # Pastikan tipe data numerik
    data['IPK'] = pd.to_numeric(data['IPK'], errors='coerce')
    data['Lama Studi (Semester)'] = pd.to_numeric(data['Lama Studi (Semester)'], errors='coerce')

    # ===============================
    # 3️⃣ JUMLAH WISUDAWAN PER PRODI
    # ===============================
    jumlah_per_prodi = data.groupby('Program Studi')['NIM'].count().reset_index()
    jumlah_per_prodi.columns = ['Program Studi', 'Jumlah Wisudawan']

    # ===============================
    # 4️⃣ KLASIFIKASI GRADE BERDASARKAN IPK
    # ===============================
    def tentukan_grade(ipk):
        if ipk >= 3.75:
            return 'A'
        elif ipk >= 3.50:
            return 'B+'
        elif ipk >= 3.00:
            return 'B'
        elif ipk >= 2.50:
            return 'C'
        else:
            return 'D'

    data['Grade'] = data['IPK'].apply(tentukan_grade)

    # ===============================
    # 5️⃣ KLASIFIKASI PREDIKAT KELULUSAN
    # ===============================
    def tentukan_predikat(row):
        ipk = row['IPK']
        lama = row['Lama Studi (Semester)']
        if ipk >= 3.75 and lama <= 8:
            return 'Cumlaude (Dengan Pujian)'
        elif ipk >= 3.50 and lama <= 9:
            return 'Sangat Memuaskan'
        elif ipk >= 3.00:
            return 'Memuaskan'
        else:
            return 'Cukup'

    data['Predikat Wisuda'] = data.apply(tentukan_predikat, axis=1)

    # ===============================
    # 6️⃣ OUTPUT DI TERMINAL
    # ===============================
    print("\n===============================")
    print("🎓 DATA WISUDAWAN TERIMPORT")
    print("===============================\n")
    print(data.head(10))  # tampilkan 10 data teratas

    print("\n===============================")
    print("📊 JUMLAH WISUDAWAN PER PRODI")
    print("===============================\n")
    print(jumlah_per_prodi)

    print("\n===============================")
    print("🏅 DISTRIBUSI PREDIKAT WISUDA")
    print("===============================\n")
    print(data['Predikat Wisuda'].value_counts())

    # Top 5 berdasarkan IPK tertinggi
    top5 = data.sort_values(by='IPK', ascending=False).head(5)
    print("\n===============================")
    print("🌟 5 MAHASISWA DENGAN IPK TERTINGGI")
    print("===============================\n")
    print(top5[['NIM', 'Nama Mahasiswa', 'Program Studi', 'IPK', 'Grade', 'Predikat Wisuda']])

    # ===============================
    # 7️⃣ VISUALISASI DATA
    # ===============================
    # --- Grafik Batang Jumlah Wisudawan per Prodi ---
    plt.figure(figsize=(10, 6))
    plt.bar(jumlah_per_prodi['Program Studi'], jumlah_per_prodi['Jumlah Wisudawan'], color='skyblue')
    plt.title('Jumlah Wisudawan per Program Studi')
    plt.xlabel('Program Studi')
    plt.ylabel('Jumlah Wisudawan')
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.show()

    # --- Grafik Pie Distribusi Predikat Kelulusan ---
    distribusi_predikat = data['Predikat Wisuda'].value_counts()
    plt.figure(figsize=(8, 8))
    plt.pie(distribusi_predikat, labels=distribusi_predikat.index, autopct='%1.1f%%', startangle=140, colors=plt.cm.Set3.colors)
    plt.title('Distribusi Predikat Wisuda')
    plt.show()

    # --- Grafik Perbandingan Rata-rata IPK antar Prodi (opsional) ---
    ipk_rata_per_prodi = data.groupby('Program Studi')['IPK'].mean().reset_index()
    plt.figure(figsize=(10, 6))
    plt.bar(ipk_rata_per_prodi['Program Studi'], ipk_rata_per_prodi['IPK'], color='orange')
    plt.title('Rata-rata IPK per Program Studi')
    plt.xlabel('Program Studi')
    plt.ylabel('Rata-rata IPK')
    plt.xticks(rotation=45, ha='right')
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.tight_layout()
    plt.show()

    # ===============================
    # 8️⃣ SIMPAN HASIL KE FILE EXCEL
    # ===============================
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        data.to_excel(writer, index=False, sheet_name='Data Wisudawan Lengkap')
        jumlah_per_prodi.to_excel(writer, index=False, sheet_name='Jumlah per Prodi')
        ipk_rata_per_prodi.to_excel(writer, index=False, sheet_name='Rata IPK per Prodi')
        top5.to_excel(writer, index=False, sheet_name='Top 5 IPK Tertinggi')

    print(f"\n✅ File hasil berhasil disimpan sebagai: {output_file}")

# ===============================
# PENANGANAN ERROR
# ===============================
except FileNotFoundError as e:
    print(f"❌ Error: {e}")

except KeyError as e:
    print(f"❌ Error: Kolom tidak ditemukan: {e}")

except ValueError as e:
    print(f"❌ Error: {e}")

except Exception as e:
    print(f"❌ Terjadi kesalahan tak terduga: {e}")

print("bagus")