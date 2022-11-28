#!/usr/bin/env python
# coding: utf-8

# In[1]:


import numpy as np
import os
import time
import matplotlib.pyplot as plt
from tabulate import tabulate

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


# In[4]:



def kategori_sei():
    kategori = input("Kategori SEI\n"+
                    "1. call center\n2. email\n3. twitter\n4. online chat\n"+
                    "Pilihan kategori SEI Anda? (1/2/3/4): ")
    print("")
    if kategori=='1':
        kategori='call center'
        kategori_call_center = input("Kategori Call Center\n"+
                                    "1. car insurance\n2. courier\n3. regular banking\n"+
                                    "Pilihan kategori Call Center Anda? (1/2/3): ")
        print("")
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
    elif kategori=='4':
        kategori='online chat'
        kategori_call_center = None        
    else:
        print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
        exit()
        
    return kategori, kategori_call_center

print("")
print("***PROGRAM SERVICE EXCELENCE INDEX (SEI)***")
print("")
olah_indeks_sei = input("Apakah Anda ingin melakukan olah indeks SEI? (y/n): ")
print("")
if olah_indeks_sei=='y':
    kategori, kategori_call_center = kategori_sei()
    data_tahun = int(input("Silakan masukkan tahun pembuatan data: "))
    print("")    
elif olah_indeks_sei=='n':
    pass
else:
    print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
    exit()

laporan_sei = input("Apakah Anda ingin membuat laporan PowerPoint? (y/n): ")
print("")
if laporan_sei=='n':
    pass
elif laporan_sei=='y':    
    if olah_indeks_sei=='n':
        kategori, kategori_call_center = kategori_sei()
        data_tahun = int(input("Silakan masukkan tahun pembuatan data: "))
        print("")    
    if kategori=='call center':
        file_data_mentah = 'data mentah/data mentah '+kategori+' - '+kategori_call_center+' - '+str(data_tahun)+'.xlsx'
    else:
        file_data_mentah = 'data mentah/data mentah '+kategori+' - '+str(data_tahun)+'.xlsx'
    data_mentah = pd.read_excel(file_data_mentah, engine='openpyxl')
    brands = data_mentah['Brand'].unique()
    print("Daftar Brand terdeteksi:")
    if kategori=='call center':
        brand_ivr, brand_nonivr = od.get_ivr_nonivr(data_mentah, brands)
        
    tabel_brand = pd.DataFrame(columns=['no','brand'])
    for idx_brd, brd in enumerate(brands):
        tabel_brand.loc[idx_brd, 'no'] = idx_brd+1
        tabel_brand.loc[idx_brd, 'brand'] = brd   
        
        if kategori=='call center':
            if brd in brand_ivr:
                tabel_brand.loc[idx_brd, 'keterangan'] = 'ivr'
            else:
                tabel_brand.loc[idx_brd, 'keterangan'] = 'non-ivr'                
    print(tabulate(tabel_brand, headers = tabel_brand.columns, tablefmt = 'psql', showindex=False))
    
    client = input("Nama brand client yang Anda ingin buatkan laporan (silakan pilih nomor-nya): ")
    client = brands[int(client)-1]
    print("")
    gambar_footer = input("PASTIKAN FILE GAMBAR ADA DI FOLDER 'gambar'\n"+
                          "Nama file beserta format extension-nya (.png/.jpg/.jpeg) yang akan digunakan sebagai gambar footer laporan: ")
    print("")
    gambar_footer = 'gambar/%s'%(gambar_footer)
    try:
        plt.imread(gambar_footer)
    except FileNotFoundError:
        print("Gambar tidak ditemukan")
        exit()
        
else:
    print("Pilihan yang Anda masukkan tidak tersedia, silakan coba lagi.")
    exit()


# # olah data

# In[ ]:



if olah_indeks_sei == 'y' or olah_indeks_sei == 'Y':
    print("\n")
    print("Olah indkes SEI %s sedang dalam proses ..."%(kategori.title()))
    start = time.time()

    if kategori=='call center':
        file_data_mentah = 'data mentah/data mentah '+kategori+' - '+kategori_call_center+' - '+str(data_tahun)+'.xlsx'
    else:
        file_data_mentah = 'data mentah/data mentah '+kategori+' - '+str(data_tahun)+'.xlsx'
    file_bobot = "bobot/bobot perhitungan "+kategori+".xlsx"

    data_mentah = pd.read_excel(file_data_mentah, engine='openpyxl')

    # LOAD DATA BOBOT
    daftar_bobot = pd.read_excel(file_bobot, sheet_name='indeks', engine='openpyxl')
    data_bobot = od.data_bobot(daftar_bobot)
    
    # GANTI NAMA BULAN KE NOMOR BULAN
    data_mentah.Bulan = od.ganti_bulan(data_mentah.Bulan,'nama bulan','nomor bulan')

    # DROP NULL ROW AFTER OBSERVATION MONTH
    data_mentah = od.drop_month_after(data_mentah)
    
    # DEFINE LIST BRAND
    list_brand = data_mentah['Brand'].unique()
    print("Data teridri dari %d brand"%(len(list_brand)))
    if kategori == 'call center':
        brand_ivr, brand_nonivr = od.get_ivr_nonivr(data_mentah, list_brand)
        print("Jumlah brand ivr %d, brand nonivr %d"%(len(brand_ivr), len(brand_nonivr)))
        
    # TABEL CROSSTAB
    crosstab_rata_transpose, crosstab_rata= od.crosstab_rata_transpose(data_mentah, list_brand, data_bobot)

    # DEFINE LIST BULAN
    list_bulan= pd.Series(data_mentah.Bulan.unique())
    list_bulan = list(od.ganti_bulan(list_bulan,'nomor bulan','nama bulan'))
    print("Bulan pengambilan data %s - %s"%(list_bulan[0], list_bulan[-1]))

    # Olah indeks
    if kategori=='call center':
        print("Membuat indeks subaspek ...")
        nilai_subaspek = od.nilai_sub_aspek(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose,
                                            brand_ivr, brand_nonivr)
        print("Membuat indeks aspek ...")
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_subaspek,
                                              brand_ivr, brand_nonivr)    
    elif kategori=='email':
        print("Membuat indeks aspek ...")
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, None)
    else:
        print("Membuat indeks subaspek ...")
        nilai_subaspek = od.nilai_sub_aspek(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose)
        print("Membuat indeks aspek ...")
        nilai_aspek = od.nilai_aspek_with_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_subaspek)    
    nilai_aspek_only = od.nilai_aspek_without_sub(kategori, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_aspek)
    df_aspek_pivot = od.df_aspek_pivot(nilai_aspek, nilai_aspek_only, list_bulan)
    tabel_summary_aspek = od.tabel_summary(df_aspek_pivot, list_bulan)

    if kategori=='call center':
        print("Membuat indeks sub-KPI ...")
        df_subkpi_pivot = od.df_subkpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_aspek_pivot,
                                             brand_ivr, brand_nonivr)
        print("Membuat indeks KPI ...")
        df_kpi_pivot = od.df_kpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_subkpi_pivot,
                                       brand_ivr, brand_nonivr)
        print("Membuat indeks dimensi ...")
        df_dimensi_pivot = od.df_dimensi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_kpi_pivot,
                                               brand_ivr, brand_nonivr)
        print("Mendapatkan indeks SEI akhir ...")
        df_ccsei_pivot = od.df_ccsei_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_dimensi_pivot,
                                           brand_ivr, brand_nonivr)    
    else:
        print("Membuat indeks sub-KPI ...")
        df_subkpi_pivot = od.df_subkpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_aspek_pivot)
        print("Membuat indeks KPI ...")
        df_kpi_pivot = od.df_kpi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_subkpi_pivot)
        print("Membuat indeks dimensi ...")
        df_dimensi_pivot = od.df_dimensi_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_kpi_pivot)
        print("Mendapatkan indeks SEI akhir ...")
        df_ccsei_pivot = od.df_ccsei_pivot(kategori, daftar_bobot, list_bulan, list_brand, df_dimensi_pivot)    

    tabel_summary_subkpi = od.tabel_summary(df_subkpi_pivot, list_bulan)
    tabel_summary_kpi = od.tabel_summary(df_kpi_pivot, list_bulan)
    tabel_summary_dimensi = od.tabel_summary(df_dimensi_pivot, list_bulan)

    # Olah indeks engagement
    bobot_engagement = pd.read_excel(file_bobot, sheet_name='engagement', engine='openpyxl')
    if kategori=='call center':
        print("Membuat indeks engagement ...")
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand, 
                                                         None,  brand_ivr, brand_nonivr)
    elif kategori=='twitter' or kategori=='online chat':
        print("Membuat indeks engagement ...")
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand)   
    else:
        print("Membuat indeks engagement ...")
        tabel_be_avg, df_eng_final_pivot = od.engagement(kategori, bobot_engagement, df_aspek_pivot, list_bulan, list_brand, 
                                                         crosstab_rata_transpose)
    # Olah indeks area perbaikan
    bobot_perbaikan = pd.read_excel(file_bobot, sheet_name='perbaikan', engine='openpyxl')
    if kategori=='call center':
        print("Membuat area perbaikan ...")
        area_perbaikan = od.area_perbaikan(kategori, data_mentah, bobot_perbaikan, list_bulan, list_brand,
                                           brand_ivr, brand_nonivr)
    else:
        print("Membuat area perbaikan ...")
        area_perbaikan = od.area_perbaikan(kategori, data_mentah, bobot_perbaikan, list_bulan, list_brand)   
        
    # print("Seluruh proses pembuatan olah indeks telah selesai")

    # simpan data olah indeks
    if kategori=='call center':
        nama_output = "tabel output/tabel output "+kategori+" - "+kategori_call_center+' - '+str(data_tahun)+".xlsx"
    else:
        nama_output = "tabel output/tabel output "+kategori+' - '+str(data_tahun)+".xlsx"
    list_simpan = [data_mentah, crosstab_rata, crosstab_rata_transpose, tabel_summary_aspek, tabel_summary_subkpi, tabel_summary_kpi,
                  tabel_summary_dimensi, df_ccsei_pivot, tabel_be_avg,
                  df_eng_final_pivot, area_perbaikan]

    list_nama_simpan = ['data mentah','crosstab rata', 'crosstab transpose', 'aspek','subkpi','kpi','dimensi','index akhir',
                       'engagement index aspek','engagement index final','aspek perbaikan']
    
    print("Menyimpan data olah indeks ...")
    writer = pd.ExcelWriter(nama_output)
    for i in range(len(list_simpan)):
        if (list_nama_simpan[i]=='data mentah'):
            list_simpan[i].to_excel(writer,sheet_name=list_nama_simpan[i], index=False, freeze_panes=(1, 3))
            worksheet = writer.sheets[list_nama_simpan[i]]
            worksheet.autofilter(0, 0, 0, 2) #first row, first col, last row, last col
        elif (list_nama_simpan[i]=='crosstab rata'):
            list_simpan[i].to_excel(writer,sheet_name=list_nama_simpan[i], freeze_panes=(6,2))        
            worksheet = writer.sheets[list_nama_simpan[i]]
            worksheet.autofilter(5, 1, 5, 1)
        elif (list_nama_simpan[i]=='crosstab transpose'):
            list_simpan[i].set_index('KODE').to_excel(writer,sheet_name=list_nama_simpan[i], freeze_panes=(3,1))
        else:
            max_col = len(list_simpan[i].index.names)-1
            list_simpan[i].reset_index().to_excel(writer,sheet_name=list_nama_simpan[i], index=False, freeze_panes=(1, max_col+1) , merge_cells=False)
            worksheet = writer.sheets[list_nama_simpan[i]]
            worksheet.autofilter(0, 0, 0, max_col)
    writer.save()
    print("")
    print("Data indeks berhasil disimpan di folder '%s' dengan nama '%s'"%(nama_output.split('/')[0],nama_output.split('/')[1]))

    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("Durasi pengolahan data: {:0>2} jam {:0>2} menit {:05.2f} detik".format(int(hours),int(minutes),seconds))

elif olah_indeks_sei == 'n' or olah_indeks_sei == 'N':
    print("")
    print("Olah indeks data tidak dilakukan")
else:
    pass


# In[ ]:


# if kategori=='call center':
#     print("brand ivr")
#     print(brand_ivr)
#     print("----------------")
#     print("brand nonivr")
#     print(brand_nonivr)


# # laporan ppt

# In[ ]:



if laporan_sei == 'y' or laporan_sei == 'Y':
    print("\n")
    print("Pembuatan laporan SEI %s sedang dalam proses ..."%(kategori.title()))
    start = time.time()

    if kategori=='email':
        bobot = 'bobot/bobot perhitungan email.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks', engine='openpyxl')
        # print("email")
        data_xls = pd.ExcelFile('tabel output/tabel output email - %s.xlsx'%(str(data_tahun)), engine='openpyxl')
    elif kategori=='call center':
        bobot = 'bobot/bobot perhitungan call center.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks', engine='openpyxl')
        # print("call center")
        data_xls = pd.ExcelFile('tabel output/tabel output call center - %s - %s.xlsx'%(kategori_call_center,str(data_tahun)), engine='openpyxl')
    elif kategori=='twitter':
        bobot = 'bobot/bobot perhitungan twitter.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks', engine='openpyxl')
        # print("twitter")
        data_xls = pd.ExcelFile('tabel output/tabel output twitter - %s.xlsx'%(str(data_tahun)), engine='openpyxl')
    elif kategori=='online chat':
        bobot = 'bobot/bobot perhitungan online chat.xlsx'
        bobot_indeks = pd.read_excel(bobot, sheet_name = 'indeks', engine='openpyxl')
        # print("online chat")
        data_xls = pd.ExcelFile('tabel output/tabel output online chat - %s.xlsx'%(str(data_tahun)), engine='openpyxl')        

    bobot_indeks = ppt_ccsei.bobot_indeks(bobot_indeks)
    
    list_bulan = pd.read_excel(data_xls, sheet_name='index akhir', engine='openpyxl').columns[1:-1].to_list()
    print("Bulan pengambilan data %s - %s"%(list_bulan[0], list_bulan[-1]))
    
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
        print("Membuat slide dimensi %s ..."%(dimensi_))
        ppt_ccsei.slide_transisi(template_ppt, gambar_footer,'KINERJA DIMENSI '+dimensi_)

        kpi_in_dimensi = bobot_indeks[bobot_indeks['DIMENSI'].isin([dimensi_])]['KPI'].unique()
        for j in range(len(kpi_in_dimensi)):
            kpi_ = kpi_in_dimensi[j]
            print("Membuat grafik tracking KPI %s ..."%(kpi_))
            ppt_ccsei.plot_grafik_tracking(template_ppt, gambar_footer, list_bulan,
                                           data_xls, dimensi_, kpi_, client)

            subkpi_in_dimensi = bobot_indeks[bobot_indeks['KPI'].isin([kpi_])]['SUB KPI'].unique()
            for k in range(len(subkpi_in_dimensi)):
                subkpi_ = subkpi_in_dimensi[k]
                if (kategori=='twitter') or (kategori=='online chat') or (kategori=='email'):
                    print("Membuat tabel data semester sub-KPI %s ..."%(subkpi_))
                    ppt_ccsei.plot_tabel_data_semester(template_ppt, gambar_footer, list_bulan, data_xls,
                                                       dimensi_, subkpi_,client, kategori, bobot_indeks)
                else:
                    print("Membuat tabel data semester sub-KPI %s ..."%(subkpi_))
                    ppt_ccsei.plot_tabel_data_semester(template_ppt, gambar_footer, list_bulan, data_xls,
                                                       dimensi_, subkpi_,client)
                print("Membuat grafik batang sub-KPI %s ..."%(subkpi_))
                ppt_ccsei.plot_barchart_tabel(template_ppt, gambar_footer, list_bulan, data_xls,
                                              dimensi_, subkpi_,client, data_tahun, kategori, bobot_indeks)

    print("Membuat slide indeks engagement ...")
    ppt_ccsei.slide_transisi(template_ppt, gambar_footer, 'Engagement Index')
    ppt_ccsei.plot_tabel_engagement(template_ppt, gambar_footer, list_bulan, data_xls, client, data_tahun)

    print("Membuat slide area perbaikan ...")
    ppt_ccsei.slide_transisi(template_ppt, gambar_footer, 'Area Perbaikan')
    if kategori=='call center':
        ppt_ccsei.plot_tabel_perbaikan(template_ppt, gambar_footer, list_bulan,
                                       data_xls, bobot, client, data_tahun, kategori)
    else:
        ppt_ccsei.plot_tabel_perbaikan(template_ppt, gambar_footer, list_bulan,
                                   data_xls, bobot, client, data_tahun, kategori)

    print("Menambahkan halaman slide ...")
    ppt_ccsei.halaman_slide(template_ppt)
    if kategori=='call center':
        nama_laporan = 'laporan/Laporan '+kategori+" "+kategori_call_center+" "+client+" "+str(data_tahun)+'.pptx'
    else:
        nama_laporan = 'laporan/Laporan '+kategori+" "+client+" "+str(data_tahun)+'.pptx'
    print("Menyimpan laporan ppt ...")
    template_ppt.save(nama_laporan)
    
    print("")
    print("Laporan untuk %s berhasil dibuat"%(client))
    print("Laporan berhasil disimpan di folder '%s' dengan nama '%s'"%(nama_laporan.split('/')[0],nama_laporan.split('/')[1]))

    end = time.time()
    hours, rem = divmod(end-start, 3600)
    minutes, seconds = divmod(rem, 60)
    print("Durasi pembuatan laporan: {:0>2} jam {:0>2} menit {:05.2f} detik".format(int(hours),int(minutes),seconds))

elif laporan_sei == 'n' or laporan_sei == 'N':
    print("")
    print("Laporan tidak dibuat")
else:
    pass


# In[ ]:


print("")
print("***Program telah selesai***")
print("")


# # end
