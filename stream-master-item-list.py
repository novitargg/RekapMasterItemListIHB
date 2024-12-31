#import library
import pandas as pd
import streamlit as st
from io import BytesIO

def app():

    st.image("logo_ps.jpeg", use_container_width=True)
    title = """
    <h1 style="text-align: center;
        white-space: nowrap;
        background-color: grey;
        padding: 10px;">
        UPDATED MASTER ITEM LIST IHB 2024
    </h1>
    """
    st.markdown(title, unsafe_allow_html=True)

    st.markdown('__________')

    #upload file Item LIst
    file_item_list = st.file_uploader("Upload File Item List Terbaru:", type ='xlsx')
    if file_item_list is not None:  # Jika file ada
        try:
            psm = pd.read_excel(file_item_list, sheet_name='PSM')
            season = pd.read_excel(file_item_list, sheet_name='Season')
            smbu = pd.read_excel(file_item_list, sheet_name='SMBU')
        except Exception as e:
            st.error(f"Ada yang salah pada file Item List yaitu : {e}")
    else:
        st.warning("Unggah File Item List Terbaru")
    
    #upload Master Item List IHB 2024
    file_tarikan_gr_all_brand = st.file_uploader("Upload File Tarikan GR All Brand Terbaru:", type ='xlsx')

    if file_tarikan_gr_all_brand is not None:
        try:
            gr_tarikan = pd.read_excel(file_tarikan_gr_all_brand, sheet_name = 'GR STORE')
            gr_tarikan = gr_tarikan.dropna(axis=1,how='all' )
            gr_tarikan = gr_tarikan.iloc[3:].reset_index(drop=True)
            gr_tarikan.columns = ['Item Code 1', 'GR 1','Item Code 2', 'GR 2','Item Code 3', 'GR 3']
            #ubah format data tarikan
            gr_tarikan['GR 2']= pd.to_datetime(gr_tarikan['GR 2']).dt.strftime('%d/%m/%Y')
        except  Exception as e:
            st.error(f'Ada yang salah pada file Tarikan GR All Brand yaitu:{e}')
    else:
        st.warning("Unggah Tarikan GR All Brand Terakhir")

    #upload Master Item List IHB 2024
    file_master_item_list_terakhir = st.file_uploader("Upload File Master Item List Terbaru:", type ='xlsx')
    if file_master_item_list_terakhir is not None:  # Jika file ada
        try:
            #read all sheet
            master_PSM = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'PSM')
            master_GR_Store = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'GR STORE')
            master_IHB = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'IHB')
            master_Store = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'MASTER STORE')
            master_Category = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'CATEGORY NEW')
            master_Range = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'RANGE')
            master_Lebaran = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'DATE LEBARAN')
            master_Season = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'SEASON')
            master_Month = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'Month')
            master_Week = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'Week')
            master_Designer = pd.read_excel(file_master_item_list_terakhir, sheet_name= 'Designer')
        except Exception as e:
            st.error(f"Ada Kesalahan Format File Master Item List IHB 2024 : {e}")

        try:
            #create new df
            add_psm = pd.DataFrame() 
            #Item No. from ket jual
            add_psm['Item No.'] = psm[pd.isna(psm['Ket Jual'])]['Item No.']

            #add variabel Item Generic, Color, dan Size
            def split_item_code(item_code):
                item_code= item_code.replace('--','-')
                parts = item_code.split('-')
                #memastikan ada 3 parts
                if len(parts) < 3:
                    parts.extend([''] * (3 - len(parts)))
                return parts
            #split Item Code
            add_psm[['Item Generic', 'Color', 'Size']] = add_psm['Item No.'].apply(split_item_code).apply(pd.Series)

            #add Item SKU
            add_psm['Item SKU'] = add_psm.apply(
                lambda row:row['Item Generic']
                if pd.isnull(row['Size']) or len(row['Size'])==0
                else row['Item Generic']+"-" + row['Color'], axis=1)
                
            #lookup by Item No.
            lookup_data_psm = (
                psm.groupby('Item No.')[['Item Description','Bar Code','Item Group','Manufacturer','Inventory UoM']].first().to_dict())
            #add Item Description
            add_psm['Item Description']= add_psm['Item No.'].map(lookup_data_psm ['Item Description'])
            #add barcode
            add_psm['Bar Code']= add_psm['Item No.'].map(lookup_data_psm ['Bar Code'])
            #add Item Group
            add_psm['Item Group']= add_psm['Item No.'].map(lookup_data_psm ['Item Group'])
            #add Manufacturer
            add_psm['Manufacturer']= add_psm['Item No.'].map(lookup_data_psm ['Manufacturer'])
            #add Inventory UoM
            add_psm['Inventory UoM']= add_psm['Item No.'].map(lookup_data_psm ['Inventory UoM'])
            
            lookup_price = (
                psm.groupby('Item No.')[['Last Eval. Price','Last Purchase Price','NormalPrice','OriginalPrice', 'PromoPrice','PurchasePrice','WholesalePrice']].first().to_dict())

            add_psm['Last Eval. Price']= add_psm['Item No.'].map(lookup_price['Last Eval. Price']).fillna(0)
            add_psm['Last Purchase Price']= add_psm['Item No.'].map(lookup_price['Last Purchase Price']).fillna(0)
            add_psm['NormalPrice']= add_psm['Item No.'].map(lookup_price['NormalPrice']).fillna(0)
            add_psm['Original Price']= add_psm['Item No.'].map(lookup_price['OriginalPrice']).fillna(0)
            add_psm['PromoPrice']= add_psm['Item No.'].map(lookup_price['PromoPrice']).fillna(0)
            add_psm['PurchasePrice']= add_psm['Item No.'].map(lookup_price['PurchasePrice']).fillna(0)
            add_psm['WholesalePrice']= add_psm['Item No.'].map(lookup_price['WholesalePrice']).fillna(0)

            #add Description Generic
            df_item_description_first = (
                add_psm.groupby('Item Generic')['Item Description'].first().to_dict())
            add_psm['Description Generic']= add_psm['Item Generic'].map(df_item_description_first)
            
            #add Description SKU
            add_psm['Description SKU'] = add_psm.apply(
                lambda row: row['Description Generic']
                if pd.isnull(row['Size']) or len(row['Size'])==0
                else row['Item Description'] [:len(row['Item Description'])-len(row['Size'])], axis=1)
        
            #add brand 
            add_psm['Brand']= add_psm['Item Group'].str.slice(6,9)

            #add Ket Brand
            lookup_ket_brand = (
                master_PSM.groupby('Brand')['Ket Brand'].first().to_dict())
            def ket_brand(row):
                try:
                    if row['Brand'] in ['DRG','INS','JEC','PSC','SPY','VOX']:
                        return 'IHB'
                    else:
                        return lookup_ket_brand.get(row['Brand'], 'ASSET')
                except:
                    return 'ASSET'
            add_psm['Ket Brand'] = add_psm.apply(ket_brand, axis=1)

            #add brand name
            lookup_brand_name = (
                master_PSM.groupby('Brand')['Brand Name'].first().to_dict())
            def name_brand(row):
                return lookup_brand_name.get(row['Brand'], 'ASSET')
            add_psm['Brand Name'] = add_psm.apply(name_brand, axis=1)

            #add DateGR
            date_now = pd.Timestamp.now()
            lookup_dategr = (
                master_GR_Store.groupby('Item SKU')['Min of GR Awal'].first().to_dict())
        
            def date_gr (row):
                return lookup_dategr.get(row['Item SKU'], date_now.date())
            add_psm['Date GR'] = add_psm.apply(date_gr, axis=1)
            
            #add ket barang
            def ket_barang(row):
                if row['Inventory UoM']=='EA':
                    return 'JUAL'
                else:
                    return "GWP / ASSET"
            add_psm['Ket Barang']= add_psm.apply(ket_barang,axis=1)
        
            #add Status
            lookup_status = (
                master_IHB.groupby('Item Code')['Status'].first().to_dict())
            def status_sales (row):
                return lookup_status.get(row['Item No.'], 'OSB')
            add_psm['Status'] = add_psm.apply(status_sales, axis=1)

            #add Season
            lookup_season_1 = (
                season.groupby('Item No.')['Season'].first().to_dict())
            lookup_season_2 = (
                master_IHB.groupby('Item Code')['Season'].first().to_dict())
            def season_psm (row):
                item_no = str(row['Item No.'])
                season = lookup_season_1.get(item_no)
                if season is not None:
                    return season
                return lookup_season_2.get(item_no, '-')
            add_psm['Season'] = add_psm.apply(season_psm, axis=1)

            #add Flag
            def flag_number(value):
                try:
                    return int(str(value)[-1])
                except (TypeError, ValueError):
                    return 'ASSET'
            add_psm['Flag']= add_psm['Item Group'].apply(flag_number)


            #sort No
            lookup_no=(
                psm.groupby('Item No.')['No'].first().to_dict())
            master_PSM['No'] = master_PSM['Item No.'].map(lookup_no)

        except Exception as e:
            st.error(f'Ada Kesalahan {e} saat melengkapi data add_psm, periksa data item list dan Master Item List IHB')


        try:
            #create add_ihb
            add_ihb = pd.DataFrame()
        
            #item No by ket IHB and brand IHB
            item_no_ihb = psm[pd.isna(psm['Ket IHB'])& psm['Brand'].isin(['DRG','INS','PSC','JEC','SPY','VOX'])][['Item No.','Brand']]
            add_ihb['Item Code']=item_no_ihb["Item No."]
        except Exception as e:
            st.error(f"Terdapat error {e} saat vlookup Item No., periksa kembali data item list")

        try:
            #makesure item code is str
            add_ihb['Item Code']=add_ihb["Item Code"].astype(str)

            #split item code
            def split_item_code_ihb(item_code_ihb):
                item_code_ihb =item_code_ihb.replace("--", "-")
                parts = item_code_ihb.split("-")
                #makesure parts have 3 parts
                if len(parts)<3:
                    parts.extend(['']*(3-len(parts)))
                return parts
            add_ihb [["Item Generic", "Color", "Size"]]= add_ihb['Item Code'].apply(split_item_code_ihb).apply(pd.Series)
            #check data
            add_ihb['Color']=add_ihb["Color"].fillna('').astype(str)
            add_ihb['Size']=add_ihb["Size"].fillna("").astype(str)
        except Exception as e:
            st.error(f"Terdapat error {e} pada saat split Item No, periksa kembali data item lis")

        try:
            
            #add item SKU
            add_ihb['Item SKU'] = add_ihb.apply(
                lambda row: row['Item Generic'] if not row['Color'] else row['Item Generic'] + "-" + row["Color"], axis=1)
            #Item Description
            add_ihb['Item Description']= add_ihb['Item Code'].map(lookup_data_psm ['Item Description'])
            #Bar Code
            add_ihb['Bar Code']= add_ihb['Item Code'].map(lookup_data_psm ['Bar Code'])
            #Item Group
            add_ihb['Item Group']= add_ihb['Item Code'].map(lookup_data_psm ['Item Group'])
            #Manufacturer
            add_ihb['Manufacturer']= add_ihb['Item Code'].map(lookup_data_psm ['Manufacturer'])
            #Brand
            add_ihb['Brand']= add_ihb['Item Group'].str.slice(6,9)
            #Inventory UoM
            add_ihb['Inventory UoM']= add_ihb['Item Code'].map(lookup_data_psm ['Inventory UoM'])
        
            #Unit Current Price
            add_ihb['Unit Current Price']= add_ihb['Item Code'].map(lookup_price['NormalPrice'])

            #Unit Cost Price from sheet SMBU 
            lookup_cost_price =(
            smbu.groupby('Item No.')['PurchasePrice'].first().to_dict())
            add_ihb['Unit Cost Price']= add_ihb['Item Code'].map(lookup_cost_price).fillna(0)
            add_ihb['Unit Whole Price']= add_ihb['Item Code'].map(lookup_price['WholesalePrice']).fillna(0)
            add_ihb['Unit Original Price']= add_ihb['Item Code'].map(lookup_price['OriginalPrice']).fillna(0)
            #description generic
            add_ihb['Description Generic']= add_ihb['Item Generic'].map(df_item_description_first)

        except Exception as e:
            st.error(f"Error{e}, saat vlookup variabel add_ihb (Item Description-Description Generic)")

        try:
            #add Description SKU
            add_ihb['Description SKU'] = add_ihb.apply(
                lambda row: row['Description Generic'] if not row['Color'] else row['Description Generic']+"-"+row['Color'], axis=1)
        except Exception as e:
            st.error(f"Kesalahan {e} saat melengkapi deskription SKU")
        try:
            #Terminology
            lookup_terminology = (
                master_Category.groupby('Type')['Main Terminology'].first().to_dict())
            type_bantu = add_ihb['Item Group'].str.slice(3,6)
            add_ihb['Terminology'] = type_bantu.map(lookup_terminology)

            #Project
            add_ihb['Project']= 'REGULER'
            #Status
            add_ihb['Status']= 'NORMAL'
            #Season
            def season_ihb (row):
                item_no = str(row['Item Code'])
                season = lookup_season_1.get(item_no)
                if season is not None:
                    return season
                return lookup_season_2.get(item_no, '-')
            add_ihb['Season'] = add_ihb.apply(season_ihb, axis=1)
            #Date Now
            
            add_ihb['Date Now'] = pd.Timestamp.now()
            add_ihb['Date Now'] = pd.to_datetime(add_ihb['Date Now'], errors='coerce')

            #GR Awal
            lookup_gr_awal = (
                gr_tarikan.groupby('Item Code 2')['GR 2'].first().to_dict())
            lookup_gr_awal_2 =(
                add_ihb.groupby('Item Code')['Date Now'].first().to_dict())
            def gr_awal(row):
                item_code_ihb = str(row['Item Code'])
                gr_awal = lookup_gr_awal.get(item_code_ihb)
                if gr_awal is not None:
                    return gr_awal
                return lookup_gr_awal_2.get(item_code_ihb, add_ihb['Date Now'])
            add_ihb['GR Awal'] = add_ihb.apply(gr_awal, axis=1)
            #GR Updated
            add_ihb['GR Updated'] = add_ihb['GR Awal']
            #GR Store
            lookup_gr_store = (
                master_GR_Store.groupby('Item SKU')['Min of GR Store'].first().to_dict())
            def gr_store(row):
                item_sku = str(row['Item SKU'])
                gr_store = lookup_gr_store.get(item_sku)
                if gr_store is not None:
                    return gr_store
                return row['Date Now']
            add_ihb['GR Store']= add_ihb.apply(gr_store, axis=1)
            add_ihb['GR Store']= pd.to_datetime(add_ihb['GR Store'], errors='coerce')
            #Ket Promo
            add_ihb['Ket Promo']= add_ihb.apply(
                lambda row: 'Promo' if row['Unit Current Price'] < row['Unit Original Price'] else 'Normal', axis=1)
            add_ihb['Aging Month'] = (add_ihb['Date Now'].dt.year - add_ihb['GR Store'].dt.year)*12 + (add_ihb['Date Now'].dt.month - add_ihb['GR Store'].dt.month)
            add_ihb['Week'] = (add_ihb['Date Now'].dt.year - add_ihb['GR Store'].dt.year)*52 + (add_ihb['Date Now'].dt.isocalendar().week - add_ihb['GR Store'].dt.isocalendar().week)

            #sub terminology
            lookup_sub_terminology = (
                master_Category.groupby('Type')['Type Name'].first().to_dict())
            type_bantu = add_ihb['Item Group'].str.slice(3,6)
            add_ihb['Sub Terminology']= type_bantu.map(lookup_sub_terminology)
            #gender
            lookup_gender = (
                master_Category.groupby('Gender')['Gender Name'].first().to_dict())
            type_bantu = add_ihb['Item Group'].str[0]
            add_ihb['Gender']= type_bantu.map(lookup_gender)
        
            add_ihb['Family Series']= "Reguler"
            add_ihb['New Status'] ="Normal Baru"

        except Exception as e:
            st.error(f"Terdapat error {e} pada saat melengkapi add_ihb, periksa kembali data item list")
        
        try:
            def excel_file_dl(df, sheet_name):
                if df.empty:
                    raise ValueError(f"Dataframe untuk {sheet_name} kosong")
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                output.seek(0)
                return output#download data add  
            psm_excel = excel_file_dl(add_psm, 'PSM')
            ihb_excel = excel_file_dl(add_ihb, 'IHB')
            st.write("ITEM YANG AKAN DITAMBAHKAN KE MASTER ITEM LIST IHB SHEET 'PSM' ")
            st.dataframe(add_psm)
            st.download_button(
                label= "Download Add PSM ",
                data = psm_excel,
                file_name = 'add_psm.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            

            st.write("ITEM YANG AKAN DITAMBAHKAN KE MASTER ITEM LIST IHB SHEET 'IHB' ")
            st.dataframe(add_ihb)
            st.download_button(
                label= "Download Add IHB ",
                data = ihb_excel,
                file_name = 'add_ihb.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception as e:
            st.error(f'Ada Kesalahan {e}, file tidak bisa di download') 
    else:
        st.warning("Unggah Master Item List IHB 2024 Terakhir")

app()


