import pandas as pd
import numpy as np
import os


"""
Program ini berisi fungsi-fungsi untuk membuat data indeks SEI (v2)
"""


"""
nama_bulan dan nomor_bulan digunakan sebagai acuan ketika konversi bulan diperlukan
"""
nomor_bulan = ['1','2','3','4','5','6','7','8','9','10','11','12']
nama_bulan = ['Januari','Februari','Maret',
              'April','Mei','Juni',
              'Juli','Agustus','September',
              'Oktober','November','Desember']


"""
daftar subkpi dalam komponen IVR
"""
subkpi_ivr = ['Kelengkapan Menu IVR', 'Kehandalan Menu IVR',
              'Ketanggapan Menu IVR', 'Kemudahan Dalam Menggunakan IVR',
              'Salam Pembuka IVR', 'Salam Penutup IVR', 'Kenyamanan IVR']


"""
daftar kolom call center yang merupakan komponen IVR
"""
kolom_ivr = ['IVR_bhs_IND','IVR_bhs_ENG',
             'aksesIVR_user','aksesIVR_nonuser',
             'menuIVR_terakses_semua','menuIVR_terakses_1_tdk','menuIVR_terakses_2_sd_3_tdk','menuIVR_terakses_4_sd_5_tdk','menuIVR_terakses_lbh_5_tdk',
             'info_produk_byIVR_semua','info_produk_byIVR_1_tdk','info_produk_byIVR_2_sd_3_tdk','info_produk_byIVR_4_sd_5_tdk','info_produk_byIVR_lbh_5_tdk',
             'menuIVR_sesuai_semua','menuIVR_sesuai_1_tdk','menuIVR_sesuai_2_sd_3_tdk','menuIVR_sesuai_4_sd_5_tdk','menuIVR_sesuai_lbh_5_tdk',
             'infoIVR_sesuai_semua','infoIVR_sesuai_1_tdk','infoIVR_sesuai_2_sd_3_tdk','infoIVR_sesuai_4_sd_5_tdk','infoIVR_sesuai_lbh_5_tdk',
             'konfirm_byIVR_semua','konfirm_byIVR_1LayTdk','konfirm_byIVR_2LayTdk','konfirm_byIVR_3LayTdk','konfirm_byIVR_semuaLayTdk',
             'responIVR_noInput_bbrpCSO','responIVR_noInput_infoCSO','responIVR_noInput_antri','responIVR_noInput_infoPutus','responIVR_noInput_putus',
             'susunan_menuIVR',
             'IVR_menuCSO_semua','IVR_menuCSO_1LayTdk','IVR_menuCSO_2LayTdk','IVR_menuCSO_3LayTdk','IVR_menuCSO_semuaLayTdk',
             'IVR_mainMenu_semua','IVR_mainMenu_1LayTdk','IVR_mainMenu_2LayTdk','IVR_mainMenu_3LayTdk','IVR_mainMenu_semuaLayTdk',
             'IVR_prevMenu_semua','IVR_prevMenu_1LayTdk','IVR_prevMenu_2LayTdk','IVR_prevMenu_3LayTdk','IVR_prevMenu_semuaLayTdk',
             'IVR_repInfo_semua','IVR_repInfo_1LayTdk','IVR_repInfo_2LayTdk','IVR_repInfo_3LayTdk','IVR_repInfo_semuaLayTdk',
             'IVR_endCall',
             'bypass_IVR',
             'pembukaIVR_waktu','pembukaIVR_welcome','pembukaIVR_brand','pembukaIVR_terimKasih',
             'penutupIVR_brand','penutupIVR_terimKasih','penutupIVR_waktu',
             'IVR_noOverlap_semua','IVR_noOverlap_1_tdk','IVR_noOverlap_2_sd_3_tdk','IVR_noOverlap_4_sd_5_tdk','IVR_noOverlap_lbh_5_tdk',
             'IVR_tdk_kemresek','IVR_tdk_dengung',
             'intonasi_IVR',
             'IVR_tdk_trlCepat','IVR_tdk_trlLambat','IVR_tdk_trlKeras','IVR_tdk_trlPelan',
             'artikulasi_IVR',
             'pelafalan_IVR',
             'bahasaIVR_MagicWords','bahasaIVR_tdkMenyalahkan','bahasaIVR_tdkMenyudutkan','bahasaIVR_tdkMemerintah',
             'tidak_ada_Iklan']


def ganti_bulan(data_series_bulan, diganti, pengganti,
                nomor_bulan=nomor_bulan, nama_bulan=nama_bulan):
    """
    Melakukan konversi nama bulan ke nomor bulan atau sebaliknya

    Parameter
    ---------
    data_series_bulan : data series bulan yang akan dikonversi
    diganti : diisi 'nomor bulan' atau 'nama bulan' sebagai parameter yang akan diganti
    pengganti : diisi 'nomor bulan' atau 'nama bulan' sebagai parameter penggantinya
    nomor_bulan : list nomor bulan
    nama_bulan : list nama-nama bulan

    Return
    ------
    data_series_bulan : data yang sudah dikonversi
    """
    if diganti=='nomor bulan':
        diganti=nomor_bulan
        pengganti=nama_bulan
    elif diganti=='nama bulan':
        diganti=nama_bulan
        pengganti=nomor_bulan
#    for i in range(len(data_series_bulan.unique())):
#        data_series_bulan.replace(diganti[i],pengganti[i], inplace=True)
    dict_bulan = dict(zip(diganti, pengganti))
    data_series_bulan = data_series_bulan.map(dict_bulan)
    return data_series_bulan


def drop_month_after(data_mentah):
    """
    Menghapus baris yang isi kolom penilaiannya kosong semua

    Parameter
    ---------
    data_mentah : dataframe yang akan diolah

    Return
    ------
    data_mentah : dataframe yang sudah tidak memiliki baris kosong di semua kolom penilaiannya
    """
    isi_data = data_mentah.iloc[:,3:]
    last_month = data_mentah.dropna(subset=data_mentah.iloc[:,3:].columns, axis=0, how='all')['Bulan'].unique()[-1]
    id_last_month = nomor_bulan.index(last_month)
    month_to_be_dropped = nomor_bulan[id_last_month+1:]
    data_mentah = data_mentah[~data_mentah['Bulan'].isin(month_to_be_dropped)]
    return data_mentah



def get_ivr_nonivr(data_mentah, list_brand):
    """
    Mendefinisikan brand yang memiliki IVR atau tidak di kategori call center

    Parameter
    ---------
    list_brand : daftar brand yang ada di data

    Return
    ------
    brand_ivr : daftar brand yang memiliki IVR
    brand_nonivr : daftar brand yang tidak memiliki IVR
    """

    brand_ivr = []
    brand_nonivr = []
    for brd in list_brand:
        kondisi = data_mentah[data_mentah['Brand']==brd][kolom_ivr].notnull().any().any()
        if kondisi==True:
            brand_ivr.append(brd)
        else:
            brand_nonivr.append(brd)

    return brand_ivr, brand_nonivr


def sum_product(value, bobot):
    """
    Melakukan sumproduct nilai dengan bobot

    Parameter
    ---------
    value : series nilai data
    bobot : bobot perkalian sesuai dimensi/kpi/subkpi/aspek/subaspek/rincian

    Return
    ------
    nilai : nilai hasil sumproduct
    """
    value_kali = value
    bobot_kali = bobot
    nilai = np.nansum(value_kali*bobot_kali)
    if value_kali.isnull().all()==True and int(nilai)==0:
        nilai = np.nan
    return nilai


def crosstab_rata_transpose(data_mentah, list_brand, daftar_bobot,
                    nomor_bulan=nomor_bulan,nama_bulan=nama_bulan):
    """
    Membuat data per brand menjadi rata-rata tiap bulan

    Parameter
    ---------
    data_mentah : dataframe yang akan diolah
    list_brand : list array brand dari dataframe
    nomor_bulan : daftar nomor bulan
    nama_bulan : daftar nama-nama bulan

    Return
    ------
    crosstab_rata_transpose : dataframe yang sudah jadi berisi rata-rata penilaian KPI per brand tiap bulan
    """

    # ADD ROW RATA-RATA DAN JUMLAH PER BULAN
    data_ratajumlah = pd.DataFrame(columns=data_mentah.columns)
    for brand in list_brand:
        temp = data_mentah[data_mentah.Brand.isin([brand])]
        for bln in temp.Bulan.unique():
            temp_perbulan = temp[temp.Bulan.isin([bln])].reset_index(drop=True)

            temp_avg = temp_perbulan.iloc[:,3:].mean().to_frame().transpose()
            temp_avg.insert(loc=0, column='Brand', value=brand)
            temp_avg.insert(loc=1, column='Bulan', value=bln)
            temp_avg.insert(loc=2, column='Periode', value='Rata-rata')

            temp_sum = temp_perbulan.iloc[:,3:].sum().to_frame().transpose()
            temp_sum.insert(loc=0, column='Brand', value=brand)
            temp_sum.insert(loc=1, column='Bulan', value=bln)
            temp_sum.insert(loc=2, column='Periode', value='Jumlah')

            temp_end = pd.concat([temp_perbulan,temp_avg, temp_sum], axis=0)

            data_ratajumlah = pd.concat([data_ratajumlah, temp_end], axis=0)
    data_ratajumlah = data_ratajumlah.reset_index(drop=True)
    data_ratajumlah.Bulan = ganti_bulan(data_ratajumlah.Bulan,nomor_bulan,nama_bulan)

    # ADD RATA-RATA INDUSTRI
    data_ratajumlah = data_ratajumlah.groupby(by=['Bulan','Brand','Periode']).mean()
    data_ratajumlah = data_ratajumlah.index.set_levels(nama_bulan, level=0)
    data_rata_groupby_bulan = data_mentah.loc[:, data_mentah.columns != 'Periode'].groupby(by=['Bulan','Brand']).mean().reset_index()
    data_rata_industri = pd.DataFrame(columns=data_rata_groupby_bulan.columns)
    for bln in data_rata_groupby_bulan.Bulan.unique():
        temp_perbulan = data_rata_groupby_bulan[data_rata_groupby_bulan.Bulan.isin([bln])].reset_index(drop=True)

        temp_ind = temp_perbulan.iloc[:,2:].mean().to_frame().transpose()
        temp_ind.insert(loc=0, column='Bulan', value=bln)
        temp_ind.insert(loc=1, column='Brand', value='~Industri')
        temp_end = pd.concat([temp_perbulan, temp_ind],axis=0)
        data_rata_industri = pd.concat([data_rata_industri, temp_end], axis=0)

    data_rata_industri = data_rata_industri.reset_index(drop=True)
    data_rata_industri.iloc[:,2:] = round(data_rata_industri.iloc[:,2:]*100,1)
    

    # CROSSTAB RATA TRANSPOSE
    crosstab_rata = data_rata_industri.groupby(by=['Bulan','Brand']).mean()
    crosstab_rata.index = crosstab_rata.index.set_levels(nama_bulan, level=0)

    kolom_dim_ = daftar_bobot[['DIMENSI','KPI','SUB KPI','ASPEK','KODE']]
    kolom_dim_ = kolom_dim_.set_index(kolom_dim_.columns.to_list()).transpose().columns
    crosstab_rata.columns = kolom_dim_
    crosstab_rata_transpose = crosstab_rata.transpose().droplevel([0,1,2,3]).reset_index()
    return crosstab_rata_transpose, crosstab_rata



def data_bobot(daftar_bobot):
    """
    Melakukan pra-pengolahan data bobot. Mengisi kolom-kolom yang kosong di bagian 'NOMOR LEVEL 1', 'DIMENSI',
    'KPI', 'SUB KPI', dan 'ASPEK'.

    Parameter
    ---------
    daftar_bobot : dataframe yang berisi daftar bobot

    Return
    ------
    daftar_bobot : dataframe bobot yang sudah tidak memiliki cell kosong di kolom yang ditentukan.
    """
    daftar_bobot['NOMOR LEVEL 1'] = daftar_bobot['NOMOR LEVEL 1'].fillna(method='ffill').astype(int)
    daftar_bobot['DIMENSI'] = daftar_bobot['DIMENSI'].fillna(method='ffill')
    daftar_bobot['KPI'] = daftar_bobot['KPI'].fillna(method='ffill')
    daftar_bobot['SUB KPI'] = daftar_bobot['SUB KPI'].fillna(method='ffill')
    daftar_bobot['ASPEK'] = daftar_bobot['ASPEK'].fillna(method='ffill')
    return daftar_bobot



def nilai_sub_aspek(kriteria, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose,
                    brand_ivr=None, brand_nonivr=None,
                    nomor_bulan=nomor_bulan, nama_bulan=nama_bulan):
    """
    Melakukan penghitungan nilai rincian untuk mendapatkan nilai subaspek dari data berkriteria 'call center', 'online chat' atau 'twitter'.
    Data 'email' tidak memiliki level nilai sampai rincian.

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    crosstab_rata_transpose : dataframe yang sudah jadi berisi rata-rata penilaian KPI per brand tiap bulan
    nomor_bulan : daftar nomor bulan
    nama_bulan : daftar nama-nama bulan

    Return
    ------
    nilai_subaspek : dataframe yang berisi nilai-nilai subaspek
    """
    if kriteria == 'call center':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT RINCIAN IVR'].notnull()].fillna(method='ffill')
    elif kriteria == 'twitter' or kriteria == 'online chat':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT RINCIAN'].notnull()].fillna(method='ffill')

    key_group = tb_bobot['NOMOR LEVEL 2'].unique()
    nilai_subaspek = []
    for bln in list_bulan:
        for nm_file in list_brand:
            for i in range(len(key_group)):
                id_key = tb_bobot[tb_bobot['NOMOR LEVEL 2'].isin([key_group[i]])].index.values

                if kriteria == 'call center':
                    if nm_file in brand_nonivr:
                        bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT RINCIAN NON-IVR']
                        value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key),(bln,nm_file)]
                        nilai = sum_product(value_kali, bobot_kali)
                        nilai = [tb_bobot['NOMOR LEVEL 1'].values[0], key_group[i], bln, nm_file, round(nilai,1)]
                    else:
                        bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT RINCIAN IVR']
                        value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key),(bln,nm_file)]
                        nilai = sum_product(value_kali, bobot_kali)
                        nilai = [tb_bobot['NOMOR LEVEL 1'].values[0], key_group[i], bln, nm_file, round(nilai,1)]
                    nilai_subaspek.append(nilai)

                elif kriteria == 'twitter' or kriteria == 'online chat':
                    bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT RINCIAN']
                    value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key),(bln,nm_file)]
                    nilai = sum_product(value_kali, bobot_kali)
                    nilai = [tb_bobot['NOMOR LEVEL 1'].values[0], key_group[i], bln, nm_file, round(nilai,1)]
                    nilai_subaspek.append(nilai)

    nilai_subaspek = pd.DataFrame(nilai_subaspek, columns=['NOMOR LEVEL 1','NOMOR LEVEL 2','Bulan','Brand','Nilai Subaspek'])
    nilai_subaspek.Bulan = ganti_bulan(nilai_subaspek.Bulan,nama_bulan,nomor_bulan)
    nilai_subaspek = nilai_subaspek.pivot(index = ['NOMOR LEVEL 1','NOMOR LEVEL 2'],
                                          columns = ['Bulan','Brand'],values='Nilai Subaspek')
    for bln in nomor_bulan[:len(list_bulan)]:
        nilai_subaspek[(bln,'~Industri')] = round(nilai_subaspek[(bln)].mean(axis=1),1)
    nilai_subaspek = nilai_subaspek.stack(dropna=False).unstack()
    nilai_subaspek.columns = nilai_subaspek.columns.set_levels(nama_bulan, level=0)
    nilai_subaspek = nilai_subaspek.reset_index()

    return nilai_subaspek



def nilai_aspek_with_sub(kriteria, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_subaspek,
                        brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai aspek dari nilai-nilai subaspek. Nilai subaspek didapat dari nilai_subaspek dan crosstab_rata_transpose.

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    crosstab_rata_transpose : data dari dataframe ini akan digunakan sebagai nilai subaspek jika kategori SEI tidak punya nilai subaspek hasil perhitungan dari rincian
    nilai_subaspek : data nilai subaspek hasil perhitungan dari rincian

    Return
    ------
    nilai_aspek : dataframe yang berisi nilai-nilai aspek hasil dari subaspek
    """
    if kriteria == 'call center':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT SUB ASPEK NON-IVR'].notnull()].fillna(method='ffill')
    elif kriteria == 'twitter' or kriteria == 'online chat':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT SUB ASPEK'].notnull()].fillna(method='ffill')
    elif kriteria == 'email':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT SUB ASPEK'].notnull()]
    key_group = tb_bobot['NOMOR LEVEL 1'].unique()
    nilai_aspek = []
    for bln in list_bulan:
        for nm_file in list_brand:
            for i in range(len(key_group)):

                # CALL CENTER ---------------------
                if kriteria == 'call center':
                    if key_group[i]==37:
                        id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                        if nm_file in brand_nonivr:
                            bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK NON-IVR'].reset_index(drop=True)
                            value_kali = nilai_subaspek.loc[:, (bln,nm_file)]
                            nilai = sum_product(value_kali, bobot_kali)
                        else:
                            bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK IVR'].reset_index(drop=True)
                            value_kali = nilai_subaspek.loc[:, (bln,nm_file)]
                            nilai = sum_product(value_kali, bobot_kali)
                    elif key_group[i]==62:
                        id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values

                        i_62 = crosstab_rata_transpose.loc[min(id_key):max(id_key)+6, (bln,nm_file)]
                        i_62a = i_62.loc[min(id_key):min(id_key)].values[0]
                        i_62b = max(i_62.loc[min(id_key)+1:min(id_key)+1].values)
                        i_62 = pd.Series([i_62a, i_62b])

                        if nm_file in brand_nonivr:
                            bobot_kali =  tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK NON-IVR'].reset_index(drop=True)
                            value_kali = i_62
                            nilai = sum_product(value_kali, bobot_kali)
                        else:
                            bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK IVR'].reset_index(drop=True)
                            value_kali = i_62
                            nilai = sum_product(value_kali, bobot_kali)
                    else:
                        id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                        if nm_file in brand_nonivr:
                            bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK NON-IVR']
                            value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key), (bln,nm_file)]
                            nilai = sum_product(value_kali, bobot_kali)
                        else:
                            bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK IVR']
                            value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key), (bln,nm_file)]
                            nilai = sum_product(value_kali, bobot_kali)
                    nilai = round(nilai,1)

                # TWITTER & ONLINE CHAT---------------------
                elif kriteria == 'twitter' or kriteria == 'online chat':
                    if key_group[i]==37:
                        id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                        ab = pd.concat([crosstab_rata_transpose.loc[46:46,(bln,nm_file)],
                                        nilai_subaspek.loc[:, (bln,nm_file)]],axis=0)
                        ab = ab.reset_index(drop=True)
                        bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK'].reset_index(drop=True)
                        value_kali = ab
                        nilai = sum_product(value_kali, bobot_kali)
                    else:
                        id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                        bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK']
                        value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key), (bln,nm_file)]
                        nilai = sum_product(value_kali, bobot_kali)
                    nilai = round(nilai,1)

                 # EMAIL ---------------------
                elif kriteria == 'email':
                    id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                    bobot_kali = tb_bobot.loc[min(id_key):max(id_key),'BOBOT SUB ASPEK']
                    value_kali = crosstab_rata_transpose.loc[min(id_key):max(id_key), (bln,nm_file)]
                    nilai = sum_product(value_kali, bobot_kali)
                    nilai = round(nilai,1)

                nilai = [tb_bobot['DIMENSI'][id_key].values[0],tb_bobot['KPI'][id_key].values[0],
                         tb_bobot['SUB KPI'][id_key].values[0],
                         key_group[i], str(key_group[i]), tb_bobot['ASPEK'][id_key].values[0],bln,nm_file,nilai]
                nilai_aspek.append(nilai)

    nilai_aspek = pd.DataFrame(nilai_aspek,
                               columns=['Dimensi','KPI','SUB KPI','Nomor','Kode','Aspek','Bulan','Brand','Nilai Aspek'])
    return nilai_aspek



def nilai_aspek_without_sub(kriteria, daftar_bobot, list_bulan, list_brand, crosstab_rata_transpose, nilai_aspek):
    """
    Melakukan penghitungan nilai aspek dari nilai-nilai subaspek. Nilai subaspek didapat dari nilai_subaspek dan crosstab_rata_transpose.

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    crosstab_rata_transpose : data dari dataframe ini akan digunakan sebagai nilai subaspek jika kategori SEI tidak punya nilai subaspek hasil perhitungan dari rincian
    nilai_aspek : dataframe nilai aspek yang didapat dari perhitunga subaspek, digunakan untuk filter nomor aspek penilaian yang belum ada

    Return
    ------
    nilai_aspek_only : dataframe yang berisi nilai-nilai aspek langsung dari crosstab_rata_transpose
    """
    sub_nomor = ['a','b','c','d','e','f']
    tb_bobot = daftar_bobot[daftar_bobot['SUB ASPEK'].isnull()]
    key_group = tb_bobot['NOMOR LEVEL 1'].unique()

    nilai_aspek_only = []
    for bln in list_bulan:
        for nm_file in list_brand:
            for i in range(len(key_group)):
                if key_group[i] in nilai_aspek['Nomor'].unique():
                    continue
                else:
                    id_key = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin([key_group[i]])].index.values
                    df_temp = crosstab_rata_transpose.loc[min(id_key):max(id_key), (bln)]
                    if len(df_temp)==1:
                        nao = crosstab_rata_transpose.loc[id_key, (bln,nm_file)].values[0]

                        if kriteria == 'call center':
                            if key_group[i]==60 or key_group[i]==61:
                                nao = nao/100
                                nao = 1-nao
                                nao = nao*100
                                nao = round(nao, 1)
                            else:
                                nao = nao
                                nao = round(nao, 1)
                            nilai = [tb_bobot['DIMENSI'][id_key].values[0],
                                     tb_bobot['KPI'][id_key].values[0],tb_bobot['SUB KPI'][id_key].values[0],
                                     int(key_group[i]), str(key_group[i]), tb_bobot['ASPEK'][id_key].values[0], bln, nm_file,nao]
                            nilai_aspek_only.append(nilai)

                        elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
                            nao = round(nao, 1)
                            nilai = [tb_bobot['DIMENSI'][id_key].values[0],tb_bobot['KPI'][id_key].values[0],
                                     tb_bobot['SUB KPI'][id_key].values[0], int(key_group[i]), str(key_group[i]),
                                     tb_bobot['ASPEK'][id_key].values[0], bln, nm_file,nao]
                            nilai_aspek_only.append(nilai)
                    else:
                        for j in range(len(df_temp)):
                            nao = crosstab_rata_transpose.loc[id_key[j],
                                                     (bln,nm_file)]
                            nao = round(nao,1)
                            nilai = [tb_bobot['DIMENSI'][id_key].values[0],tb_bobot['KPI'][id_key].values[0],
                                     tb_bobot['SUB KPI'][id_key].values[0], int(key_group[i]), str(key_group[i])+"_"+sub_nomor[j],
                                     tb_bobot['ASPEK'][id_key].values[j], bln,nm_file,nao]
                            nilai_aspek_only.append(nilai)

    nilai_aspek_only = pd.DataFrame(nilai_aspek_only,
                                    columns=['Dimensi','KPI','SUB KPI','Nomor','Kode','Aspek','Bulan','Brand','Nilai Aspek'])
    return nilai_aspek_only



def df_aspek_pivot(nilai_aspek, nilai_aspek_only, list_bulan):
    """
    Menggabungkan nilai aspek dari perhitungan subaspek dan nilai aspek dari perhitungan rata-rata biasa.

    Parameter
    ---------
    nilai_aspek: dataframe nilai aspek dari perhitungan subaspek
    nilai_aspek_only : dataframe nilai aspek dari perhitungan rata-rata data mentah
    list_bulan : list array bulan yang ada di dalam dataframe

    Return
    ------
    df_aspek_pivot : dataframe yang berisi nilai aspek per brand tiap bulan
    """
    df_aspek = pd.concat([nilai_aspek, nilai_aspek_only], axis=0)
    df_aspek = df_aspek.sort_values(by=['Nomor','Kode'])
    df_aspek = df_aspek.reset_index(drop=True)

    for i in df_aspek.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_aspek.Bulan = df_aspek.Bulan.replace(i,no_bln)
    df_aspek = df_aspek.sort_values(by=['Nomor','Kode','Bulan'])
    df_aspek = df_aspek.reset_index(drop=True)

    df_aspek_pivot = df_aspek.pivot(index=['Dimensi','KPI','SUB KPI','Nomor','Kode','Aspek'],
                                    columns=['Bulan','Brand'],values='Nilai Aspek')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_aspek_pivot[(bln,'~Industri')] = round(df_aspek_pivot[(bln)].mean(axis=1),1)
    df_aspek_pivot = df_aspek_pivot.stack(dropna=False).unstack()
    df_aspek_pivot.columns = df_aspek_pivot.columns.set_levels(nama_bulan, level=0)
    df_aspek_pivot = df_aspek_pivot.sort_index(level=[3,4])

    return df_aspek_pivot



def tabel_summary(df_pivot, list_bulan):
    """
    Mengonversi tabel data dua kolom (bulan dan brand) menjadi kolom satu level (bulan)

    Parameter
    ---------
    df_pivot : dataframe yang akan diolah, kolom dataframe terdiri dua level (bulan dan brand)
    list_bulan : daftar bulan yang ada di data mentah

    Return
    ------
    df_pivot_avg : dataframe hasil dengan satu level kolom (bulan saja)
    """
    df_pivot_avg = df_pivot.copy()
    df_pivot_avg.columns = df_pivot_avg.columns.set_levels(nomor_bulan, level=0)
    df_pivot_avg = df_pivot_avg.stack(dropna=False)
    df_pivot_avg.columns = nama_bulan[:len(list_bulan)]
    df_pivot_avg['~Average'] = round(df_pivot_avg.mean(axis=1),1)

    return df_pivot_avg




def df_subkpi_pivot(kriteria, daftar_bobot, list_bulan, list_brand, df_aspek_pivot,
                   brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai subkpi dari nilai-nilai aspek

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    df_aspek_pivot : dataframe nilai aspek

    Return
    ------
    df_subkpi_pivot : dataframe yang berisi nilai-nilai subkpi
    """
    if kriteria == 'call center':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT ASPEK IVR'].notnull()]
    elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT ASPEK'].notnull()]
    key_group = tb_bobot['SUB KPI'].unique()

    nilai_subkpi = []

    for bln in list_bulan:
        for nm_file in list_brand:
            for i in key_group:
                id_key = list(tb_bobot[tb_bobot['SUB KPI'].isin([i])]['NOMOR LEVEL 1'].unique())

                aspek_ = df_aspek_pivot[df_aspek_pivot.index.get_level_values('Nomor').isin(list(id_key))]
                aspek_ = aspek_.loc[:,(bln,nm_file)].reset_index()[(bln,nm_file)]
                aspek_ = aspek_.reset_index(drop = True)

                if kriteria == 'call center':
                    if nm_file in brand_nonivr:
                        bobot_kali = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin(id_key)]['BOBOT ASPEK NON-IVR'].reset_index(drop=True)
                        value_kali = aspek_
                        nilai_subkpi_ = sum_product(value_kali, bobot_kali)
                    else:
                        bobot_kali = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin(id_key)]['BOBOT ASPEK IVR'].reset_index(drop=True)
                        value_kali = aspek_
                        nilai_subkpi_ = sum_product(value_kali, bobot_kali)

                elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
                    bobot_kali = tb_bobot[tb_bobot['NOMOR LEVEL 1'].isin(id_key)]['BOBOT ASPEK'].reset_index(drop=True)
                    value_kali = aspek_
                    nilai_subkpi_ = sum_product(value_kali, bobot_kali)

                nilai_subkpi_ = round(nilai_subkpi_,1)
                subkpi_ = [tb_bobot[tb_bobot['SUB KPI'].isin([i])]['DIMENSI'].values[0],
                           tb_bobot[tb_bobot['SUB KPI'].isin([i])]['KPI'].values[0],
                           tb_bobot[tb_bobot['SUB KPI'].isin([i])]['SUB KPI'].values[0],
                           bln, nm_file,nilai_subkpi_]

                nilai_subkpi.append(subkpi_)

    df_subkpi = pd.DataFrame(nilai_subkpi, columns=['Dimensi','KPI','SUB KPI','Bulan','Brand','Nilai Sub KPI'])

    for i in df_subkpi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_subkpi.Bulan = df_subkpi.Bulan.replace(i,no_bln)
    df_subkpi = df_subkpi.reset_index(drop=True)

    df_subkpi_pivot = df_subkpi.pivot(index=['Dimensi','KPI','SUB KPI'], columns=['Bulan','Brand'],values='Nilai Sub KPI')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_subkpi_pivot[(bln,'~Industri')] = round(df_subkpi_pivot[(bln)].mean(axis=1),1)
    df_subkpi_pivot = df_subkpi_pivot.stack(dropna=False).unstack()
    df_subkpi_pivot.columns = df_subkpi_pivot.columns.set_levels(nama_bulan, level=0)

    return df_subkpi_pivot



def df_kpi_pivot(kriteria, daftar_bobot, list_bulan, list_brand, df_subkpi_pivot,
                brand_ivr=None, brand_nonivr=None):

    """
    Melakukan penghitungan nilai kpi dari nilai-nilai subkpi

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    df_subkpi_pivot : dataframe nilai subkpi

    Return
    ------
    df_kpi_pivot : dataframe yang berisi nilai-nilai kpi
    """

    tb_bobot = daftar_bobot[daftar_bobot['SUB KPI'].notnull()]
    key_group = tb_bobot['KPI'].unique()
    nilai_kpi = []

    for bln in list_bulan:
        for nm_file in list_brand:
            i_kpi = df_subkpi_pivot.reset_index()
            for i in key_group:

                if kriteria == 'call center':
                    if i=='Accessibility' or i=='Availability' or i=='Connection Speed':
                        nilai_kpi_ = i_kpi[i_kpi['SUB KPI']==i][(bln,nm_file)].values[0]
                    else:
                        temp_bobot = tb_bobot[tb_bobot['KPI'].isin([i])]
                        daftar_sub = list(temp_bobot['SUB KPI'].unique())

                        if nm_file in brand_nonivr:
                            temp_bobot = temp_bobot[temp_bobot['BOBOT SUB KPI NON-IVR'].notnull()].sort_values(by='SUB KPI')
                            temp_data = i_kpi[i_kpi['SUB KPI'].isin(daftar_sub)] .sort_values(by='SUB KPI')
                            bobot_kali = temp_bobot['BOBOT SUB KPI NON-IVR'].reset_index(drop=True)
                            value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                            nilai_kpi_ = sum_product(value_kali, bobot_kali)
                        else:
                            temp_bobot = temp_bobot[temp_bobot['BOBOT SUB KPI IVR'].notnull()].sort_values(by='SUB KPI')
                            temp_data = i_kpi[i_kpi['SUB KPI'].isin(daftar_sub)].sort_values(by='SUB KPI')
                            bobot_kali = temp_bobot['BOBOT SUB KPI IVR'].reset_index(drop=True)
                            value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                            nilai_kpi_ = sum_product(value_kali, bobot_kali)

                else:
                    if ((kriteria == 'twitter' or kriteria=='online chat') and (i=='Probing' or i=='Providing Solution' or i=='Closing')) or ((kriteria=='email') and (i=='Feasibility' or i=='Accessibility' or i=='Availability')):
                        nilai_kpi_ = i_kpi[i_kpi['SUB KPI']==i][(bln,nm_file)].values[0]
                    else:
                        temp_bobot = tb_bobot[tb_bobot['KPI'].isin([i])]
                        daftar_sub = list(temp_bobot['SUB KPI'].unique())

                        temp_bobot = temp_bobot[temp_bobot['BOBOT SUB KPI'].notnull()].sort_values(by='SUB KPI')
                        temp_data = i_kpi[i_kpi['SUB KPI'].isin(daftar_sub)].sort_values(by='SUB KPI')
                        bobot_kali = temp_bobot['BOBOT SUB KPI'].reset_index(drop=True)
                        value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                        nilai_kpi_ = sum_product(value_kali, bobot_kali)

                nilai_kpi_ = round(nilai_kpi_,1)
                kpi_ = [tb_bobot[tb_bobot['KPI'].isin([i])]['DIMENSI'].values[0], i,bln,nm_file, nilai_kpi_]
                nilai_kpi.append(kpi_)

    df_kpi = pd.DataFrame(nilai_kpi, columns=['Dimensi','KPI','Bulan','Brand','Nilai KPI'])

    for i in df_kpi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_kpi.Bulan = df_kpi.Bulan.replace(i,no_bln)
    df_kpi = df_kpi.reset_index(drop=True)

    df_kpi_pivot = df_kpi.pivot(index=['Dimensi','KPI'], columns=['Bulan','Brand'],values='Nilai KPI')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_kpi_pivot[(bln,'~Industri')] = round(df_kpi_pivot[(bln)].mean(axis=1),1)
    df_kpi_pivot = df_kpi_pivot.stack(dropna=False).unstack()
    df_kpi_pivot.columns = df_kpi_pivot.columns.set_levels(nama_bulan, level=0)
    return df_kpi_pivot




def df_dimensi_pivot(kriteria, daftar_bobot, list_bulan, list_brand, df_kpi_pivot,
                    brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai dimensi dari nilai-nilai kpi

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    df_kpi_pivot : dataframe nilai kpi

    Return
    ------
    df_dimensi_pivot : dataframe yang berisi nilai-nilai dimensi
    """
    if kriteria == 'call center':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT KPI IVR'].notnull()]
    elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT KPI'].notnull()]
    key_group = tb_bobot['DIMENSI'].unique()

    nilai_dimensi = []
    i_dimensi = df_kpi_pivot.reset_index()
    for bln in list_bulan:
        for nm_file in list_brand:
            for i in key_group:
                temp_bobot = tb_bobot[tb_bobot['DIMENSI'].isin([i])]
                daftar_sub = temp_bobot['KPI'].unique()

                if kriteria == 'call center':
                    if nm_file in brand_nonivr:
                        temp_bobot = temp_bobot[temp_bobot['BOBOT KPI NON-IVR'].notnull()].sort_values(by='KPI')
                        temp_data = i_dimensi[i_dimensi['KPI'].isin(daftar_sub)].sort_values(by='KPI')
                        bobot_kali = temp_bobot['BOBOT KPI NON-IVR'].reset_index(drop=True)
                        value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                        nilai_dimensi_ = sum_product(value_kali, bobot_kali)
                    else:
                        temp_bobot = temp_bobot[temp_bobot['BOBOT KPI IVR'].notnull()].sort_values(by='KPI')
                        temp_data = i_dimensi[i_dimensi['KPI'].isin(daftar_sub)].sort_values(by='KPI')
                        bobot_kali = temp_bobot['BOBOT KPI IVR'].reset_index(drop=True)
                        value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                        nilai_dimensi_ = sum_product(value_kali, bobot_kali)

                elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
                    temp_bobot = temp_bobot[temp_bobot['BOBOT KPI'].notnull()].sort_values(by='KPI')
                    temp_data = i_dimensi[i_dimensi['KPI'].isin(daftar_sub)].sort_values(by='KPI')
                    bobot_kali = temp_bobot['BOBOT KPI'].reset_index(drop=True)
                    value_kali = temp_data[(bln,nm_file)].reset_index(drop=True)
                    nilai_dimensi_ = sum_product(value_kali, bobot_kali)

                nilai_dimensi_ = round(nilai_dimensi_,1)
                dimensi_ = [i,bln,nm_file, nilai_dimensi_]
                nilai_dimensi.append(dimensi_)

    df_dimensi = pd.DataFrame(nilai_dimensi, columns=['Dimensi','Bulan','Brand','Nilai Dimensi'])

    for i in df_dimensi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_dimensi.Bulan = df_dimensi.Bulan.replace(i,no_bln)
    df_dimensi = df_dimensi.reset_index(drop=True)

    df_dimensi_pivot = df_dimensi.pivot(index=['Dimensi'], columns=['Bulan','Brand'],values='Nilai Dimensi')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_dimensi_pivot[(bln,'~Industri')] = round(df_dimensi_pivot[(bln)].mean(axis=1),1)
    df_dimensi_pivot = df_dimensi_pivot.stack(dropna=False).unstack()
    df_dimensi_pivot.columns = df_dimensi_pivot.columns.set_levels(nama_bulan, level=0)

    return df_dimensi_pivot




def df_ccsei_pivot(kriteria, daftar_bobot, list_bulan, list_brand, df_dimensi_pivot,
                  brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai indeks akhir SEI dari nilai-nilai dimensi

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    daftar_bobot : dataframe bobot
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    df_dimensi_pivot : dataframe nilai dimensi

    Return
    ------
    df_ccsei_pivot : dataframe yang berisi nilai-nilai SEI akhir
    """

    if kriteria == 'call center':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT DIMENSI IVR'].notnull()]
        tb_bobot = tb_bobot[tb_bobot['BOBOT DIMENSI IVR'].notnull()].sort_values(by='DIMENSI')
    elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
        tb_bobot = daftar_bobot[daftar_bobot['BOBOT DIMENSI'].notnull()]
        tb_bobot = tb_bobot[tb_bobot['BOBOT DIMENSI'].notnull()].sort_values(by='DIMENSI')
    ccsei = []
    for bln in list_bulan:
        for nm_file in list_brand:
            i_indeks = df_dimensi_pivot.reset_index()[(bln,nm_file)]

            if kriteria == 'call center':
                if nm_file in brand_ivr:
                    bobot_kali = tb_bobot['BOBOT DIMENSI IVR'].reset_index(drop=True)
                    value_kali = i_indeks.reset_index(drop=True)
                    indeks_final = sum_product(value_kali, bobot_kali)
                else:
                    bobot_kali = tb_bobot['BOBOT DIMENSI NON-IVR'].reset_index(drop=True)
                    value_kali = i_indeks.reset_index(drop=True)
                    indeks_final = sum_product(value_kali, bobot_kali)

            elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
                bobot_kali = tb_bobot['BOBOT DIMENSI'].reset_index(drop=True)
                value_kali = i_indeks.reset_index(drop=True)
                indeks_final = sum_product(value_kali, bobot_kali)

            indeks_final = round(indeks_final,1)
            ccsei.append([bln, nm_file, indeks_final])

    df_ccsei = pd.DataFrame(ccsei, columns=['Bulan','Brand','Indeks'])

    for i in df_ccsei.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_ccsei.Bulan = df_ccsei.Bulan.replace(i,no_bln)
    df_ccsei = df_ccsei.reset_index(drop=True)

    df_ccsei_pivot = df_ccsei.pivot(index='Brand',columns=['Bulan'], values='Indeks')
    df_ccsei_pivot = df_ccsei_pivot.transpose()
    df_ccsei_pivot['~Industri'] = round(df_ccsei_pivot.mean(axis=1),1)
    df_ccsei_pivot = df_ccsei_pivot.transpose()
    df_ccsei_pivot.columns = nama_bulan[:len(list_bulan)]
    df_ccsei_pivot['~Average'] = round(df_ccsei_pivot.mean(axis=1),1)

    return df_ccsei_pivot





def eng_calltwit(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand,
                brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai engagement khusus untuk kriteria 'call center', 'online chat' dan 'twitter'

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    bobot_engagement : dataframe bobot untuk perhitungan engagement index
    df_aspek_pivot : dataframe nilai aspek
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe

    Return
    ------
    tabel_be_avg : dataframe indeks engagement per aspek
    df_eng_final_pivot : dataframe indeks engagement keseluruhan tiap brand
    """

    bobot_engagement['DIMENSI'] = bobot_engagement['DIMENSI'].fillna(method='ffill')
    bobot_engagement['KPI'] = bobot_engagement['KPI'].fillna(method='ffill')
    be_aspek = bobot_engagement.ASPEK.values

    # BY ASPEK----------
    tabel_be = df_aspek_pivot[df_aspek_pivot.index.get_level_values('Aspek').isin(be_aspek)].copy()
    tabel_be_avg = tabel_summary(tabel_be, list_bulan)

    # BY KPI----------
    eg_kpi = []
    list_kpi = tabel_be.index.get_level_values('KPI').unique()
    for bln in list_bulan:
        for nm_file in list_brand:
            for kpi_ in list_kpi:
                tb_be = tabel_be[tabel_be.index.get_level_values('KPI').isin([kpi_])]
                val_eg = tb_be.loc[:, (bln, nm_file)].reset_index(drop=True)
                if kriteria == 'call center':
                    if nm_file in brand_ivr:
                        bbt = bobot_engagement[bobot_engagement['KPI'].isin([kpi_])]['BOBOT ASPEK IVR'].reset_index(drop=True)
                        bobot_kali = bbt
                        value_kali = val_eg
                        nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    else:
                        bbt = bobot_engagement[bobot_engagement['KPI'].isin([kpi_])]['BOBOT ASPEK NON-IVR'].reset_index(drop=True)
                        bobot_kali = bbt
                        value_kali = val_eg
                        nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    eg_kpi.append([tb_be.index.get_level_values('Dimensi').unique()[0], kpi_, bln, nm_file, nilai_ ])

                elif kriteria == 'twitter' or kriteria == 'online chat':
                    bbt = bobot_engagement[bobot_engagement['KPI'].isin([kpi_])]['BOBOT ASPEK'].reset_index(drop=True)
                    bobot_kali = bbt
                    value_kali = val_eg
                    nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    eg_kpi.append([tb_be.index.get_level_values('Dimensi').unique()[0], kpi_, bln, nm_file, nilai_ ])

    df_eg_kpi = pd.DataFrame(eg_kpi, columns = ['Dimensi','KPI','Bulan','Brand','Indeks'])

    for i in df_eg_kpi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eg_kpi.Bulan = df_eg_kpi.Bulan.replace(i,no_bln)
    df_eg_kpi = df_eg_kpi.reset_index(drop=True)

    df_eg_kpi_pivot = df_eg_kpi.pivot(index=['Dimensi','KPI'], columns=['Bulan','Brand'],values='Indeks')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_eg_kpi_pivot[(bln,'~Industri')] = round(df_eg_kpi_pivot[(bln)].mean(axis=1),1)
    df_eg_kpi_pivot = df_eg_kpi_pivot.stack(dropna=False).unstack()
    df_eg_kpi_pivot.columns = df_eg_kpi_pivot.columns.set_levels(nama_bulan, level=0)

    df_eg_kpi_pivot_avg = tabel_summary(df_eg_kpi_pivot, list_bulan)

    # BY DIMENSI----------
    if kriteria == 'call center':
        bbt = bobot_engagement[bobot_engagement['BOBOT KPI IVR'].notnull()]
    elif kriteria == 'twitter' or kriteria == 'online chat':
        bbt = bobot_engagement[bobot_engagement['BOBOT KPI'].notnull()]
    eg_dimensi = []
    list_dimensi = tabel_be.index.get_level_values('Dimensi').unique()
    for bln in list_bulan:
        for nm_file in list_brand:
            for dimensi_ in list_dimensi:
                tb_be = df_eg_kpi_pivot[df_eg_kpi_pivot.index.get_level_values('Dimensi').isin([dimensi_])]
                val_eg = tb_be.loc[:, (bln, nm_file)].reset_index(drop=True)

                if kriteria == 'call center':
                    if nm_file in brand_ivr:
                        bbt_ = bbt[bbt['DIMENSI'].isin([dimensi_])]['BOBOT KPI IVR'].reset_index(drop=True)
                        bobot_kali = bbt_
                        value_kali = val_eg
                        nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    else:
                        bbt_ = bbt[bbt['DIMENSI'].isin([dimensi_])]['BOBOT KPI NON-IVR'].reset_index(drop=True)
                        bobot_kali = bbt_
                        value_kali = val_eg
                        nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    eg_dimensi.append([dimensi_, bln, nm_file, nilai_ ])

                elif kriteria == 'twitter' or kriteria == 'online chat':
                    if dimensi_ == 'Navigating' or dimensi_ == 'Human Touching':
                        bbt_ = bbt[bbt['DIMENSI'].isin([dimensi_])].sort_values(by=['BOBOT KPI'])['BOBOT KPI'].reset_index(drop=True)
                        bobot_kali = bbt_
                    else:
                        bbt_ = bbt[bbt['DIMENSI'].isin([dimensi_])]['BOBOT KPI'].reset_index(drop=True)
                        bobot_kali = bbt_
                    value_kali = val_eg
                    nilai_ = round(sum_product(value_kali, bobot_kali),1)
                    eg_dimensi.append([dimensi_, bln, nm_file, nilai_ ])

    df_eg_dimensi = pd.DataFrame(eg_dimensi, columns = ['Dimensi','Bulan','Brand','Indeks'])

    for i in df_eg_dimensi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eg_dimensi.Bulan = df_eg_dimensi.Bulan.replace(i,no_bln)
    df_eg_dimensi = df_eg_dimensi.reset_index(drop=True)

    df_eg_dimensi_pivot = df_eg_dimensi.pivot(index=['Dimensi'], columns=['Bulan','Brand'],values='Indeks')
    for bln in nomor_bulan[:len(list_bulan)]:
        df_eg_dimensi_pivot[(bln,'~Industri')] = round(df_eg_dimensi_pivot[(bln)].mean(axis=1),1)
    df_eg_dimensi_pivot = df_eg_dimensi_pivot.stack(dropna=False).unstack()
    df_eg_dimensi_pivot.columns = df_eg_dimensi_pivot.columns.set_levels(nama_bulan, level=0)

    df_eg_dimensi_pivot_avg = tabel_summary(df_eg_dimensi_pivot, list_bulan)

    # FINAL----------
    if kriteria == 'call center':
        bbt = bobot_engagement[bobot_engagement['BOBOT DIMENSI IVR'].notnull()].sort_values(by='DIMENSI')
    elif kriteria == 'twitter' or kriteria == 'online chat':
        bbt = bobot_engagement[bobot_engagement['BOBOT DIMENSI'].notnull()].sort_values(by='DIMENSI')
    engg_final = []
    for bln in list_bulan:
        for nm_file in list_brand:
            i_eng = df_eg_dimensi_pivot.reset_index()[(bln,nm_file)]

            if kriteria == 'call center':
                if nm_file in brand_ivr:
                    bobot_kali = bbt['BOBOT DIMENSI IVR'].reset_index(drop=True)
                    value_kali = i_eng.reset_index(drop=True)
                    i_final = sum_product(value_kali, bobot_kali)
                else:
                    bobot_kali = bbt['BOBOT DIMENSI NON-IVR'].reset_index(drop=True)
                    value_kali = i_eng.reset_index(drop=True)
                    i_final = sum_product(value_kali, bobot_kali)

            elif kriteria == 'twitter' or kriteria == 'online chat':
                bobot_kali = bbt['BOBOT DIMENSI'].reset_index(drop=True)
                value_kali = i_eng.reset_index(drop=True)
                i_final = sum_product(value_kali, bobot_kali)
            i_final = round(i_final,1)
            engg_final.append([bln, nm_file, i_final])
    df_eng_final = pd.DataFrame(engg_final, columns=['Bulan','Brand','Indeks'])

    for i in df_eng_final.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eng_final.Bulan = df_eng_final.Bulan.replace(i,no_bln)
    df_eng_final = df_eng_final.reset_index(drop=True)

    df_eng_final_pivot = df_eng_final.pivot(index='Brand',columns=['Bulan'], values='Indeks')
    df_eng_final_pivot = df_eng_final_pivot.transpose()
    df_eng_final_pivot['~Industri'] = round(df_eng_final_pivot.mean(axis=1),1)
    df_eng_final_pivot = df_eng_final_pivot.transpose()
    df_eng_final_pivot.columns = nama_bulan[:len(list_bulan)]
    df_eng_final_pivot['~Average'] = round(df_eng_final_pivot.mean(axis=1),1)

    return tabel_be_avg, df_eng_final_pivot




def eng_email(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand, crosstab_rata_transpose):
    """
    Melakukan penghitungan nilai engagement khusus untuk kriteria 'email'

    Parameter
    ---------
    kriteria : kriteria SEI, khusus 'email'
    bobot_engagement : dataframe bobot untuk perhitungan engagement index
    df_aspek_pivot : dataframe nilai aspek
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    crosstab_rata_transpose : data dari dataframe ini akan digunakan sebagai nilai subaspek jika kategori SEI tidak punya nilai subaspek hasil perhitungan dari rincian

    Return
    ------
    tabel_be_complete_avg : dataframe indeks engagement per aspek
    df_eng_final_pivot : dataframe indeks engagement keseluruhan tiap brand
    """


    bobot_engagement['DIMENSI'] = bobot_engagement['DIMENSI'].fillna(method='ffill')
    bobot_engagement['KPI'] = bobot_engagement['KPI'].fillna(method='ffill')
    be_aspek = bobot_engagement.ASPEK.values

    bobot_engagement['NOMOR LEVEL 1'] = bobot_engagement['NOMOR LEVEL 1'].fillna(method='ffill')
    bobot_engagement['SUB KPI'] = bobot_engagement['SUB KPI'].fillna(method='ffill')
    bobot_engagement['ASPEK'] = bobot_engagement['ASPEK'].fillna(method='ffill')

    aspek=df_aspek_pivot.reset_index().iloc[:,4:].copy()
    aspek = aspek.drop(('Aspek',''), axis = 1)
    eng_subaspek=pd.DataFrame(columns = aspek.columns)

    arr_subaspek = []
    arr_aspek = []
    arr_subkpi = []
    arr_kpi = []
    arr_dimensi = []
    arr_nomor = []
    for i in bobot_engagement['NOMOR LEVEL 1'].unique():
        tempe = bobot_engagement[bobot_engagement['NOMOR LEVEL 1'].isin([i])]
        if i==7 or i==12:
            new_asp = tempe["ASPEK"]+" - " +tempe["SUB ASPEK"]
            new_asp = new_asp.to_list()
            temp_kode = tempe['KODE'].values
            tempe2 = crosstab_rata_transpose[crosstab_rata_transpose['KODE'].isin(temp_kode)].copy()
            tempe2 = tempe2.rename(columns = {'KODE':'Kode'})

            eng_subaspek = eng_subaspek.append(tempe2)
            arr_subaspek.extend(new_asp)

            arr_nomor.extend(tempe['NOMOR LEVEL 1'].to_list())
            arr_aspek.extend(tempe['ASPEK'].to_list())
            arr_subkpi.extend(tempe['SUB KPI'].to_list())
            arr_kpi.extend(tempe['KPI'].to_list())
            arr_dimensi.extend(tempe['DIMENSI'].to_list())

    eng_subaspek[('Nomor','')] = arr_nomor
    eng_subaspek[('Subaspek','')] = arr_subaspek
    eng_subaspek[('Aspek','')] = arr_aspek
    eng_subaspek[('SUB KPI', '')] = arr_subkpi
    eng_subaspek[('KPI','')] = arr_kpi
    eng_subaspek[('Dimensi','')] = arr_dimensi
    eng_subaspek = eng_subaspek.reset_index(drop=True)
    eng_subaspek = eng_subaspek.set_index(['Dimensi','KPI','SUB KPI','Nomor','Aspek','Kode','Subaspek'])


    # BY ASPEK
    tabel_be_uncomplete = df_aspek_pivot[df_aspek_pivot.index.get_level_values('Aspek').isin(be_aspek)]
    tabel_be_uncomplete = tabel_be_uncomplete[~(tabel_be_uncomplete.index.get_level_values('Nomor').isin([7,12]))]

    tabel_be_complete = pd.concat([tabel_be_uncomplete.reset_index(), eng_subaspek.reset_index()], axis=0, ignore_index=False)
    tabel_be_complete = tabel_be_complete.reset_index(drop=True)
    for i in tabel_be_complete['Aspek'].unique():
        tmp_asp = tabel_be_complete[tabel_be_complete['Aspek'].isin([i])]
        idx_asp = tmp_asp.index
        if str(tmp_asp['Subaspek'].unique()[0])=='nan':
            tabel_be_complete.loc[idx_asp, ('Subaspek','')] = i
    tabel_be_complete = tabel_be_complete.set_index(['Dimensi','KPI','SUB KPI','Nomor','Aspek','Kode','Subaspek'])
    tabel_be_complete = tabel_be_complete.sort_index(level=[3,5])
    tabel_be_complete = tabel_be_complete.droplevel(level=4, axis=0)
    tabel_be_complete.index = tabel_be_complete.index.rename(['Dimensi','KPI','SUB KPI','Nomor','Kode','Aspek'])
    tabel_be_complete = tabel_be_complete.reindex(columns=nama_bulan[:len(list_bulan)],level=0)

    tabel_be_complete_avg = tabel_summary(tabel_be_complete, list_bulan)

    # BY SUB KPI----------
    eg_subkpi = []
    list_subkpi = tabel_be_complete.index.get_level_values('SUB KPI').unique()

    for bln in list_bulan:
        for nm_file in list_brand:
            for list_subkpi_ in list_subkpi:
                tb_be = tabel_be_complete[tabel_be_complete.index.get_level_values('SUB KPI').isin([list_subkpi_])]
                val_eg = tb_be.loc[:, (bln, nm_file)].reset_index(drop=True)
                bbt = bobot_engagement[bobot_engagement['SUB KPI'].isin([list_subkpi_])]['BOBOT ASPEK']
                bbt = bbt[bbt.notnull()].reset_index(drop=True)
                bobot_kali = bbt
                value_kali = val_eg
                nilai_ = round(sum_product(value_kali, bobot_kali),1)
                eg_subkpi.append([tb_be.index.get_level_values('Dimensi').unique()[0], tb_be.index.get_level_values('KPI').unique()[0],
                                  list_subkpi_, bln, nm_file, nilai_ ])

    df_eg_subkpi = pd.DataFrame(eg_subkpi, columns = ['Dimensi','KPI','SUB KPI','Bulan','Brand','Indeks'])

    for i in df_eg_subkpi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eg_subkpi.Bulan = df_eg_subkpi.Bulan.replace(i,no_bln)
    df_eg_subkpi = df_eg_subkpi.reset_index(drop=True)
    df_eg_subkpi_pivot = df_eg_subkpi.pivot(index=['Dimensi','KPI','SUB KPI'], columns=['Bulan','Brand'],values='Indeks')

    for bln in nomor_bulan[:len(list_bulan)]:
        df_eg_subkpi_pivot[(bln,'~Industri')] = round(df_eg_subkpi_pivot[(bln)].mean(axis=1),1)
    df_eg_subkpi_pivot = df_eg_subkpi_pivot.stack(dropna=False).unstack()
    df_eg_subkpi_pivot.columns = df_eg_subkpi_pivot.columns.set_levels(nama_bulan, level=0)

    df_eg_subkpi_pivot_avg = tabel_summary(df_eg_subkpi_pivot, list_bulan)

    # BY KPI----------

    eg_kpi = []
    list_kpi = df_eg_subkpi_pivot.index.get_level_values('KPI').unique()

    for bln in list_bulan:
        for nm_file in list_brand:
            for kpi_ in list_kpi:
                tb_be = df_eg_subkpi_pivot[df_eg_subkpi_pivot.index.get_level_values('KPI').isin([kpi_])]
                val_eg = tb_be.loc[:, (bln, nm_file)].reset_index(drop=True)
                bbt = bobot_engagement[bobot_engagement['KPI'].isin([kpi_])]
                bbt = bbt[bbt['BOBOT SUB KPI'].notnull()]['BOBOT SUB KPI'].reset_index(drop=True)
                bobot_kali = bbt
                value_kali = val_eg
                nilai_ = round(sum_product(value_kali, bobot_kali),1)
                eg_kpi.append([tb_be.index.get_level_values('Dimensi').unique()[0], kpi_, bln, nm_file, nilai_ ])

    df_eg_kpi = pd.DataFrame(eg_kpi, columns = ['Dimensi','KPI','Bulan','Brand','Indeks'])

    for i in df_eg_kpi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eg_kpi.Bulan = df_eg_kpi.Bulan.replace(i,no_bln)
    df_eg_kpi = df_eg_kpi.reset_index(drop=True)
    df_eg_kpi_pivot = df_eg_kpi.pivot(index=['Dimensi','KPI'], columns=['Bulan','Brand'],values='Indeks')

    for bln in nomor_bulan[:len(list_bulan)]:
        df_eg_kpi_pivot[(bln,'~Industri')] = round(df_eg_kpi_pivot[(bln)].mean(axis=1),1)
    df_eg_kpi_pivot = df_eg_kpi_pivot.stack(dropna=False).unstack()
    df_eg_kpi_pivot.columns = df_eg_kpi_pivot.columns.set_levels(nama_bulan, level=0)

    df_eg_kpi_pivot_avg = tabel_summary(df_eg_kpi_pivot, list_bulan)

    # BY DIMENSI----------

    eg_dimensi = []
    list_dimensi = df_eg_kpi_pivot.index.get_level_values('Dimensi').unique()

    for bln in list_bulan:
        for nm_file in list_brand:
            for kpi_ in list_dimensi:
                tb_be = df_eg_kpi_pivot[df_eg_kpi_pivot.index.get_level_values('Dimensi').isin([kpi_])]
                val_eg = tb_be.loc[:, (bln, nm_file)].reset_index(drop=True)
                bbt = bobot_engagement[bobot_engagement['DIMENSI'].isin([kpi_])]
                bbt = bbt[bbt['BOBOT KPI'].notnull()]['BOBOT KPI'].reset_index(drop=True)
                bobot_kali = bbt
                value_kali = val_eg
                nilai_ = round(sum_product(value_kali, bobot_kali),1)
                eg_dimensi.append([kpi_, bln, nm_file, nilai_ ])

    df_eg_dimensi = pd.DataFrame(eg_dimensi, columns = ['Dimensi','Bulan','Brand','Indeks'])

    for i in df_eg_dimensi.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eg_dimensi.Bulan = df_eg_dimensi.Bulan.replace(i,no_bln)
    df_eg_dimensi = df_eg_dimensi.reset_index(drop=True)
    df_eg_dimensi_pivot = df_eg_dimensi.pivot(index=['Dimensi'], columns=['Bulan','Brand'],values='Indeks')

    for bln in nomor_bulan[:len(list_bulan)]:
        df_eg_dimensi_pivot[(bln,'~Industri')] = round(df_eg_dimensi_pivot[(bln)].mean(axis=1),1)
    df_eg_dimensi_pivot = df_eg_dimensi_pivot.stack(dropna=False).unstack()
    df_eg_dimensi_pivot.columns = df_eg_dimensi_pivot.columns.set_levels(nama_bulan, level=0)

    df_eg_dimensi_pivot_avg = tabel_summary(df_eg_dimensi_pivot, list_bulan)

    # FINAL----------

    bbt = bobot_engagement[bobot_engagement['BOBOT DIMENSI'].notnull()].sort_values(by='DIMENSI')

    engg_final = []
    for bln in list_bulan:
        for nm_file in list_brand:
            i_eng = df_eg_dimensi_pivot.reset_index()[(bln,nm_file)]
            bobot_kali = bbt['BOBOT DIMENSI'].reset_index(drop=True)
            value_kali = i_eng.reset_index(drop=True)
            i_final = sum_product(value_kali, bobot_kali)
            i_final = round(i_final,1)
            engg_final.append([bln, nm_file, i_final])
    df_eng_final = pd.DataFrame(engg_final, columns=['Bulan','Brand','Indeks'])

    for i in df_eng_final.Bulan.unique():
        idx = nama_bulan.index(i)
        no_bln = nomor_bulan[idx]
        df_eng_final.Bulan = df_eng_final.Bulan.replace(i,no_bln)
    df_eng_final.reset_index(drop=True, inplace=True)

    df_eng_final_pivot = df_eng_final.pivot(index='Brand',columns=['Bulan'], values='Indeks')
    df_eng_final_pivot = df_eng_final_pivot.transpose()
    df_eng_final_pivot['~Industri'] = round(df_eng_final_pivot.mean(axis=1),1)
    df_eng_final_pivot = df_eng_final_pivot.transpose()
    df_eng_final_pivot.columns = nama_bulan[:len(list_bulan)]
    df_eng_final_pivot['~Average'] = round(df_eng_final_pivot.mean(axis=1),1)

    return tabel_be_complete_avg, df_eng_final_pivot




def engagement(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand, crosstab_rata_transpose=None,
              brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai engagement. Fungsi ini memanggil dua fungsi lainnya sesuai dengan kriteria masukannya.

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    bobot_engagement : dataframe bobot untuk perhitungan engagement index
    df_aspek_pivot : dataframe nilai aspek
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe
    crosstab_rata_transpose : data dari dataframe ini akan digunakan sebagai nilai subaspek jika kategori SEI tidak punya nilai subaspek hasil perhitungan dari rincian, jika kriteria 'call center', 'online chat' atau 'twitter' maka tidak perlu (None)

    Return
    ------
    tabel_be_avg : dataframe indeks engagement per aspek
    df_eng_final_pivot : dataframe indeks engagement keseluruhan tiap brand
    """
    if kriteria == 'call center':
        tabel_be_avg, df_eng_final_pivot = eng_calltwit(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand,
                                                       brand_ivr, brand_nonivr)
    elif kriteria == 'twitter' or kriteria == 'online chat':
        tabel_be_avg, df_eng_final_pivot = eng_calltwit(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand)
    elif kriteria == 'email':
        tabel_be_avg, df_eng_final_pivot = eng_email(kriteria, bobot_engagement, df_aspek_pivot, list_bulan, list_brand,
                                                     crosstab_rata_transpose)
    return tabel_be_avg, df_eng_final_pivot




def area_perbaikan(kriteria, data_mentah, bobot_perbaikan, list_bulan, list_brand,
                  brand_ivr=None, brand_nonivr=None):
    """
    Melakukan penghitungan nilai indeks perbaikan

    Parameter
    ---------
    kriteria : kriteria SEI, diisi 'call center','email', 'online chat' atau 'twitter' sebagai penentu perhitungan bobot
    data_mentah : dataframe berisi data sebelum melalui proses pengolahan
    bobot_perbaikan : dataframe bobot untuk perhitungan indeks perbaikan
    list_bulan : list array bulan yang ada di dalam dataframe
    list_brand : list brand yang ada di dalam dataframe

    Return
    ------
    df_aspek_perbaikan : dataframe indeks perbaikan keseluruhan tiap brand
    """

    bobot_perbaikan['NOMOR LEVEL 1'] = bobot_perbaikan['NOMOR LEVEL 1'].fillna(method='ffill')
    bobot_perbaikan['SUB KPI'] = bobot_perbaikan['SUB KPI'].fillna(method='ffill')
    kolom_perbaikan = list(bobot_perbaikan['KODE'])

    tb_perbaikan = data_mentah.loc[:,['Brand','Bulan']+kolom_perbaikan]
    tb_perbaikan = (len(list_bulan)*len(data_mentah.Periode.unique()))-tb_perbaikan.groupby(by=['Brand']).sum().transpose()
    tb_perbaikan.index = tb_perbaikan.index.rename('KODE')
    tb_perbaikan = tb_perbaikan.reset_index()

    aspek_perbaikan = pd.concat([bobot_perbaikan[['NOMOR LEVEL 1','NOMOR LEVEL 2','KODE','SUB KPI','ASPEK LAPORAN']],
                                 tb_perbaikan.iloc[:,1:]], axis=1)
    aspek_perbaikan = aspek_perbaikan.set_index(['NOMOR LEVEL 1','NOMOR LEVEL 2','KODE','SUB KPI','ASPEK LAPORAN'])
    aspek_perbaikan.columns = pd.MultiIndex.from_product([aspek_perbaikan.columns,['Frekuensi Perbaikan']])

    if kriteria == 'call center':
        for nm_file in list_brand:
            if nm_file in brand_ivr:
                aspek_perbaikan[(nm_file,'_Priority Index')] = np.round((aspek_perbaikan[(nm_file,'Frekuensi Perbaikan')].values*bobot_perbaikan['BOBOT IMPACT IVR'].values),3)
            else:
                aspek_perbaikan[(nm_file,'_Priority Index')] = np.round((aspek_perbaikan[(nm_file,'Frekuensi Perbaikan')].values*bobot_perbaikan['BOBOT IMPACT NON-IVR'].values),3)

    elif kriteria == 'twitter' or kriteria == 'online chat' or kriteria == 'email':
        for nm_file in list_brand:
            aspek_perbaikan[(nm_file,'_Priority Index')] = np.round((aspek_perbaikan[(nm_file,'Frekuensi Perbaikan')].values*bobot_perbaikan['BOBOT IMPACT'].values),3)

    df_aspek_perbaikan = aspek_perbaikan.stack(level=0,dropna=False)
    old_idx = list(df_aspek_perbaikan.index.names)
    new_idx = old_idx[:-1]+['Brand']
    df_aspek_perbaikan.index.names = new_idx
    df_aspek_perbaikan = df_aspek_perbaikan.sort_index(level=[0,1])

    return df_aspek_perbaikan
