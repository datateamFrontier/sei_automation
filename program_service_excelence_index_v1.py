#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import os
import time

import pandas as pd
pd.options.display.max_columns = None

import ppt_ccsei
import olah_data as od


# In[2]:


"""
olah_indeks_sei = tulis 'y' atau 'Y' jika akan melakukan pengolahan indeks. Tulis 'n' atau 'N' jika tidak ingin melakukan olah indeks
laporan_sei = tulis 'y' atau 'Y' jika akan membuat laporan. Tulis 'n' atau 'N' jika tidak ingin membuat laporan

OLAH DATA
kategori : diisi 'call center' atau 'twitter' atau 'email'
kategori_call_center : hanya digunakan jika kategori 'call center' untuk menentukan brand ivr dan non-ivr. Perlu diisi dengan 'car insurance', 'courier', 'regular banking' (bisa jadi ada tambahan lain)

PEMBUATAN LAPORAN PPT
data_tahun : digunakan untuk menulis tahun di laporan
client : nama brand yang akan dibuatkan laporan
gambar_footer : logo di footer yang akan digunakan
"""


# In[3]:


# olah_indeks_sei = 'y'
# laporan_sei = 'y'

# kategori = 'call center'
# kategori_call_center = 'car insurance'
# client = 'Garda Akses Asuransi Astra'
# gambar_footer = 'gambar/Asuransi Astra2.png'

# kategori = 'email'
# client = 'customercare@commbank.co.id'
# gambar_footer='gambar/COMMBANK.jpg'

# kategori = 'twitter'
# client = '@GardaOto'
# gambar_footer = 'gambar/Asuransi Astra2.png'

# data_tahun = 2021


# In[6]:


def kategori_sei():
    kategori = input("Kategori SEI\n"+
                    "1. call center\n2. email\n3. twitter\n"+
                    "Pilihan kategori SEI Anda? (1/2/3): ")
    print('-----------------------------------------------------------------------')
    if kategori=='1':
        kategori='call center'
        kategori_call_center = input("Kategori Call Center\n"+
                                    "1. car insurance\n2. courier\n3. regular banking\n"+
                                    "Pilihan kategori Call Center Anda? (1/2/3): ")
        print('-----------------------------------------------------------------------')
        if kategori_call_center=='1':
            kategori_call_center='car insurance'
        elif kategori_call_center=='2':
            kategori_call_center='courier'
        elif kategori_call_center=='3':
            kategori_call_center='regular banking'
        else:
            print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
            exit()
    elif kategori=='2':
        kategori='email'
        kategori_call_center = None
    elif kategori=='3':
        kategori='twitter'
        kategori_call_center = None
    else:
        print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
        exit()
    return kategori, kategori_call_center

print("***PROGRAM SERVICE EXCELENCE INDEX (SEI)***")
olah_indeks_sei = input("Apakah Anda ingin melakukan olah indeks SEI? (y/n): ")
print('-----------------------------------------------------------------------')
if olah_indeks_sei=='y':
    kategori, kategori_call_center = kategori_sei()
elif olah_indeks_sei=='n':
    pass
else:
    print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
    exit()

laporan_sei = input("Apakah Anda ingin membuat laporan PowerPoint? (y/n): ")
print('-----------------------------------------------------------------------')
if laporan_sei=='n':
    pass
elif laporan_sei=='y':
    if olah_indeks_sei=='n':
        kategori, kategori_call_center = kategori_sei()
    file_data_mentah = 'data mentah/data mentah '+kategori+'.xlsx'
    data_mentah = pd.read_excel(file_data_mentah)
    brands = data_mentah['Brand'].unique()
    print("Daftar Brand terdeteksi:")
    for idx_brd, brd in enumerate(brands):
        print("%d. %s"%(idx_brd+1, brd))
    client = input("Nama brand client yang Anda ingin buatkan laporan (silakan pilih nomor-nya): ")
    client = brands[int(client)-1]
    print('-----------------------------------------------------------------------')
    gambar_footer = input("PASTIKAN FILE GAMBAR ADA DI FOLDER 'gambar'\n"+
                          "Nama file beserta format extension-nya (.png/.jpg/.jpeg) yang akan digunakan sebagai gambar footer laporan: ")
    print('-----------------------------------------------------------------------')
    gambar_footer = 'gambar/%s'%(gambar_footer)
    data_tahun = int(input("Silakan masukkan tahun pembuatan data: "))
    print('-----------------------------------------------------------------------')    
else:
    print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
    exit()

if olah_indeks_sei=='y' and laporan_sei=='n':
    data_tahun = int(input("Silakan masukkan tahun pembuatan data: "))
    print('-----------------------------------------------------------------------')


# # olah data

# In[7]:


if olah_indeks_sei == 'y' or olah_indeks_sei == 'Y':
    print("Olah indkes SEI sedang dalam proses ...")
    start = time.time()

    file_data_mentah = 'data mentah/data mentah '+kategori+'.xlsx'
    file_bobot = "bobot/bobot perhitungan "+kategori+".xlsx"

    data_mentah = pd.read_excel(file_data_mentah)

    # LOAD DATA BOBOT
    daftar_bobot = pd.read_excel(file_bobot, sheet_name='indeks')
    data_bobot = od.data_bobot(daftar_bobot)

    # GANTI NAMA BULAN KE NOMOR BULAN
    data_mentah.Bulan = od.ganti_bulan(data_mentah.Bulan,'nama bulan','nomor bulan')

    # DROP NULL ROW
    data_mentah = od.drop_null_row(data_mentah)

    # DEFINE LIST BRAND
    list_brand = data_mentah['Brand'].unique()
    if kategori == 'call center':
        brand_ivr, brand_nonivr = od.get_ivr_nonivr(list_brand, kategori_call_center)

    # TABEL CROSSTAB
    crosstab_rata_transpose= od.crosstab_rata_transpose(data_mentah, list_brand)

    # DEFINE LIST BULAN
    list_bulan= pd.Series(data_mentah.Bulan.unique())
    list_bulan = list(od.ganti_bulan(list_bulan,'nomor bulan','nama bulan'))

    # Olah indeks
    if kategori=='call center':
        nilai_subaspek = od.nilai_sub_aspek(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose,
                                           kategori_call_center)
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_subaspek,
                                             kategori_call_center)
    elif kategori=='email':
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, None)
    else:
        nilai_subaspek = od.nilai_sub_aspek(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose)
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_subaspek)
    nilai_aspek_only = od.nilai_aspek_without_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_aspek)
    df_aspek_pivot = od.df_aspek_pivot(nilai_aspek, nilai_aspek_only, list_bulan)
    tabel_summary_aspek = od.tabel_summary(df_aspek_pivot, list_bulan)

    if kategori=='call center':
        df_subkpi_pivot = od.df_subkpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_aspek_pivot,
                                            kategori_call_center)
        df_kpi_pivot = od.df_kpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_subkpi_pivot,
                                      kategori_call_center)
        df_dimensi_pivot = od.df_dimensi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_kpi_pivot,
                                              kategori_call_center)
        df_ccsei_pivot = od.df_ccsei_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_dimensi_pivot,
                                          kategori_call_center)
    else:
        df_subkpi_pivot = od.df_subkpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_aspek_pivot)
        df_kpi_pivot = od.df_kpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_subkpi_pivot)
        df_dimensi_pivot = od.df_dimensi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_kpi_pivot)
        df_ccsei_pivot = od.df_ccsei_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_dimensi_pivot)

    tabel_summary_subkpi = od.tabel_summary(df_subkpi_pivot, list_bulan)
    tabel_summary_kpi = od.tabel_summary(df_kpi_pivot, list_bulan)
    tabel_summary_dimensi = od.tabel_summary(df_dimensi_pivot, list_bulan)

    # Olah indeks engagement
    bobot_engagement = pd.read_excel(file_bobot, sheet_name='engagement')
    if kategori=='call center':
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand,
                                                         None, kategori_call_center)
    elif kategori=='twitter':
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand)
    else:
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand,
                                                         crosstab_rata_transpose)
    # Olah indeks area perbaikan
    bobot_perbaikan = pd.read_excel(file_bobot, sheet_name='perbaikan')
    if kategori=='call center':
        area_perbaikan = od.area_perbaikan(kategori, data_mentah, bobot_perbaikan, list_bulan, list_brand,
                                          kategori_call_center)
    else:
        area_perbaikan = od.area_perbaikan(kategori, data_mentah, bobot_perbaikan, list_bulan, list_brand)

    index_by_kpi, index_by_dimensi = od.tabel_indeks(tabel_summary_kpi, tabel_summary_dimensi, list_bulan)
    print("data indeks selesai dibuat")

    # simpan data olah indeks
    nama_output = "tabel output/tabel output "+kategori+" "+str(data_tahun)+".xlsx"
    list_simpan = [data_mentah, crosstab_rata_transpose, tabel_summary_aspek, tabel_summary_subkpi, tabel_summary_kpi,
                  tabel_summary_dimensi, index_by_kpi, index_by_dimensi, df_ccsei_pivot, tabel_be_avg,
                  df_eng_final_pivot, area_perbaikan]
    list_nama_simpan = ['data mentah','crosstab rata-rata','aspek','subkpi','kpi','dimensi','index by kpi','index by dimensi','index akhir',
                       'engagement index aspek','engagement index final','aspek perbaikan']
    writer = pd.ExcelWriter(nama_output)
    for i in range(len(list_simpan)):
        if (list_nama_simpan[i]=='data mentah'):
            list_simpan[i].to_excel(writer,sheet_name=list_nama_simpan[i], index=False)
        else:
            list_simpan[i].to_excel(writer,sheet_name=list_nama_simpan[i])
    writer.save()
    print("data indeks berhasil disimpan di folder '%s' dengan nama '%s'"%(nama_output.split('/')[0],nama_output.split('/')[1]))

    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("durasi pengolahan data: {:0>2} jam {:0>2} menit {:05.2f} detik".format(int(hours),int(minutes),seconds))

elif olah_indeks_sei == 'n' or olah_indeks_sei == 'N':
    print("olah indeks data tidak dilakukan")
else:
    pass


# In[ ]:





# # laporan ppt

# In[8]:


if laporan_sei == 'y' or laporan_sei == 'Y':
    print("Pembuatan laporan sedang dalam proses ...")
    start = time.time()

    if kategori=='email':
        bobot = 'bobot/bobot perhitungan email.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks')
        data_xls = pd.ExcelFile('tabel output/tabel output email %s.xlsx'%(str(data_tahun)))
    elif kategori=='call center':
        bobot = 'bobot/bobot perhitungan call center.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks')
        data_xls = pd.ExcelFile('tabel output/tabel output call center %s.xlsx'%(str(data_tahun)))
    elif kategori=='twitter':
        bobot = 'bobot/bobot perhitungan twitter.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks')
        data_xls = pd.ExcelFile('tabel output/tabel output twitter %s.xlsx'%(str(data_tahun)))

    bobot_indeks = ppt_ccsei.bobot_indeks(bobot_indeks)
    list_bulan = pd.read_excel(data_xls, sheet_name='index akhir').columns[1:-1].to_list()

    # pembuatan laporan power point
    template_ppt = ppt_ccsei.FileTemplate()

    ppt_ccsei.slide_cover(template_ppt, kategori, data_tahun, list_bulan)
    ppt_ccsei.slide_pembuka(template_ppt, gambar_footer)
    ppt_ccsei.slide_daftar_isi(template_ppt, gambar_footer, kategori)
    ppt_ccsei.slide_transisi(template_ppt, gambar_footer, 'Frame Work dan\nDefinisi Dimensi dan KPI')
    ppt_ccsei.slide_framework(template_ppt, gambar_footer, kategori)
    ppt_ccsei.slide_definisi_dimensi(template_ppt, gambar_footer, kategori)

    list_dimensi = bobot_indeks['DIMENSI'].unique()
    for i in range(len(list_dimensi)):
        dimensi_ = list_dimensi[i]
        ppt_ccsei.slide_transisi(template_ppt, gambar_footer,'KINERJA DIMENSI '+dimensi_)

        kpi_in_dimensi = bobot_indeks[bobot_indeks['DIMENSI'].isin([dimensi_])]['KPI'].unique()
        for j in range(len(kpi_in_dimensi)):
            kpi_ = kpi_in_dimensi[j]
            ppt_ccsei.plot_grafik_tracking(template_ppt, gambar_footer, list_bulan,
                                           data_xls, dimensi_, kpi_, client)

            subkpi_in_dimensi = bobot_indeks[bobot_indeks['KPI'].isin([kpi_])]['SUB KPI'].unique()
            for k in range(len(subkpi_in_dimensi)):
                subkpi_ = subkpi_in_dimensi[k]
                if (kategori=='twitter') or (kategori=='email'):
                    ppt_ccsei.plot_tabel_data_semester(template_ppt, gambar_footer, list_bulan, data_xls,
                                                       dimensi_, subkpi_,client, kategori, bobot_indeks)
                else:
                    ppt_ccsei.plot_tabel_data_semester(template_ppt, gambar_footer, list_bulan, data_xls,
                                                       dimensi_, subkpi_,client)
                ppt_ccsei.plot_barchart_tabel(template_ppt, gambar_footer, list_bulan, data_xls,
                                              dimensi_, subkpi_,client, data_tahun, kategori, bobot_indeks)

    ppt_ccsei.slide_transisi(template_ppt, gambar_footer, 'Engagement Index')
    ppt_ccsei.plot_tabel_engagement(template_ppt, gambar_footer, list_bulan, data_xls, client, data_tahun)

    ppt_ccsei.slide_transisi(template_ppt, gambar_footer, 'Area Perbaikan')
    if kategori=='call center':
        ppt_ccsei.plot_tabel_perbaikan(template_ppt, gambar_footer, list_bulan,
                                       data_xls, bobot, client, data_tahun, kategori, kategori_call_center)
    else:
        ppt_ccsei.plot_tabel_perbaikan(template_ppt, gambar_footer, list_bulan,
                                   data_xls, bobot, client, data_tahun, kategori)

    ppt_ccsei.halaman_slide(template_ppt)
    nama_laporan = 'laporan/Laporan '+kategori+" "+client+" "+str(data_tahun)+'.pptx'
    template_ppt.save(nama_laporan)
    print("laporan untuk %s berhasil dibuat"%(client))
    print("laporan berhasil disimpan di folder '%s' dengan nama '%s'"%(nama_laporan.split('/')[0],nama_laporan.split('/')[1]))

    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("durasi pembuatan laporan: {:0>2} jam {:0>2} menit {:05.2f} detik".format(int(hours),int(minutes),seconds))

elif laporan_sei == 'n' or laporan_sei == 'N':
    print("laporan tidak dibuat")
else:
    pass


# In[9]:


print("***Program telah selesai***")


# In[ ]:
