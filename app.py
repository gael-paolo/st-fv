import streamlit as st
import pandas as pd
import numpy as np

def cargar_excel(archivo, sheet_name=None, dtype=None, parts = False, usar_keep_default_na=True):
    if archivo is not None:
        try:
            if parts:
                return pd.read_excel(archivo, sheet_name=sheet_name, dtype=dtype, keep_default_na=parts)
            else:
                return pd.read_excel(archivo, sheet_name=sheet_name, dtype=dtype)
        except Exception as e:
            print(f"Error al cargar el archivo: {e}")
            return None
    return None

st.title("Aplicación Financial View Preprocess")

st.header("Análisis VN")

# Cargar archivos
archivo_dic = st.file_uploader("Sube el archivo 0_Diccionarios.xlsx", type=["xls", "xlsx"])
archivo_u = st.file_uploader("Sube el archivo U1120.xls", type=["xls", "xlsx"])
archivo_sald = st.file_uploader("Sube el archivo Saldos_VN.xls", type=["xls", "xlsx"])

if archivo_u and archivo_dic and archivo_sald:
    # Cargar dataframes
    U = cargar_excel(archivo_u, sheet_name = 'Informe', dtype={'C.comp': str})
    U['Modelo_Short'] = U['Modelo'].str[:3]
    
    Dic_VN = cargar_excel(archivo_dic, sheet_name='Models')
    Dic_Deal = cargar_excel(archivo_dic, sheet_name='Dealers', dtype={'C.comp': str})
    Dic_VN_COD = cargar_excel(archivo_dic, sheet_name='FV_VN_COD')
    Dic_VN_COD2 = cargar_excel(archivo_dic, sheet_name='FV_VN_COD2')
    
    Dic_VN.rename(columns={'CÓDIGO': "Modelo_Short", 'FAMILIA FV': "FAMILIA_FV", 'FAMILIA QT': "FAMILIA_Q"}, inplace=True)
    Dic_VN = Dic_VN[['Modelo_Short', 'FAMILIA_Q', 'FAMILIA_FV']]
    Dic_Deal = Dic_Deal[['C.comp', 'Ciudad']]
    Dic_Deal['C.comp'] = Dic_Deal['C.comp'].astype(object)
    
    Sald = cargar_excel(archivo_sald, sheet_name ='Informe', dtype = {'Saldo':'float'})
    Sald.dropna(subset=['Titulo de la cuenta'], inplace=True)
    
    def extraer_familia_Q(titulo):
        palabras = titulo.split()
        return palabras[-2] if len(palabras) > 1 else palabras[0]
    
    Sald['FAMILIA_Q'] = Sald['Titulo de la cuenta'].apply(extraer_familia_Q)
    
    # Uniones & Procesos
    U = pd.merge(U, Dic_VN, on='Modelo_Short', how='left')
    U = pd.merge(U, Dic_Deal, on='C.comp', how='left')
    
    U['Rango'] = U['Dias'].apply(lambda x: '<30' if x < 30 else '<60' if x < 60 else '<90' if x < 90 else '>90')
    U = U[['Ciudad','FAMILIA_FV', 'Rango', 'T.Costo']]
    
    Sald['C.comp'] = Sald['Cuenta mayor'].str[-4:]
    Sald = pd.merge(Sald, Dic_Deal, on='C.comp', how='left')
    Sald.drop(columns=['C.comp'], inplace=True)
    Sald = Sald[['Ciudad','FAMILIA_Q', 'Saldo']]
    
    Dic_VN = Dic_VN[['FAMILIA_Q', 'FAMILIA_FV']].drop_duplicates()
    Sald = pd.merge(Sald, Dic_VN[['FAMILIA_Q', 'FAMILIA_FV']], on='FAMILIA_Q', how='left')
    Sald.drop(columns=['FAMILIA_Q'], inplace=True)
    Sald = Sald.groupby(['Ciudad', 'FAMILIA_FV'], as_index=False)['Saldo'].sum()
    Saldo_Cont = Sald['Saldo'].sum()

    U = U.groupby(['Ciudad', 'FAMILIA_FV', 'Rango'], as_index=False)['T.Costo'].sum()
    U["Pct_Comp"] = U["T.Costo"] / U.groupby(['Ciudad', 'FAMILIA_FV'])["T.Costo"].transform("sum")
    Saldo_Inv = U['T.Costo'].sum()
    U.drop(columns=['T.Costo'], inplace=True)
    
    U = pd.merge(U, Sald, on=['Ciudad', 'FAMILIA_FV'], how='left')
    U['Comp_Saldo'] = U['Pct_Comp'] * U['Saldo']
    U.drop(columns=['Pct_Comp', 'Saldo'], inplace=True)
    df_viz = pd.pivot_table(U, values='Comp_Saldo', index = ['Ciudad', 'FAMILIA_FV'], columns='Rango', aggfunc='sum')
    df_viz = df_viz.fillna(0)
    cols_format = ['<60', '<90', '>90']
    for col in cols_format:
        df_viz[col] = df_viz[col].apply(lambda x: '{:,.2f}'.format(x))

    st.write("El análisis de stock presenta el siguiente resultado:")
    st.dataframe(df_viz)

    # Stock General
    FV_VN_1 = pd.merge(Dic_VN_COD, Sald, on=['Ciudad','FAMILIA_FV'], how='inner')[['COD_FV', 'Saldo']]
    
    # Stock por Rangos
    FV_VN_2 = pd.merge(Dic_VN_COD2, U, on=['Ciudad','FAMILIA_FV', 'Rango'], how='left')
    FV_VN_2['Comp_Saldo'] = FV_VN_2['Comp_Saldo'].fillna(0)
    FV_VN_2 = FV_VN_2[['COD_FV', 'Comp_Saldo']]
    FV_VN_2.rename(columns={"Comp_Saldo": "Saldo"}, inplace=True)
    
    # Reporte Final
    Dif_Valores = Saldo_Cont - (Saldo_Inv * 6.96)
    st.write(f"La Diferencia entre el Saldo Contable ({Saldo_Cont:,.2f} Bs.) y el Saldo Físico ({Saldo_Inv * 6.96:,.2f} Bs) es: {Dif_Valores:,.2f} Bs")

    del U, Sald, Dic_VN, Dic_Deal, Dic_VN_COD, Dic_VN_COD2, Dif_Valores, Saldo_Cont, Saldo_Inv

    FV_VN_Report = pd.concat([FV_VN_1, FV_VN_2], ignore_index=True)
    
    # Botón para descargar CSV
    csv = FV_VN_Report.to_csv(index=False).encode('utf-8')
    st.download_button(label="Descargar Reporte", data=csv, file_name="FV_VN_Report.csv", mime='text/csv')
    
    st.dataframe(FV_VN_Report)
 
st.header("Análisis Parts")

# Cargar archivos
archivo_u221 = st.file_uploader("Sube el archivo U221.xls", type=["xls", "xlsx"])
archivo_u257 = st.file_uploader("Sube el archivo U257.xls", type=["xls", "xlsx"])
archivo_sald_rep = st.file_uploader("Sube el archivo Saldos_Rep.xls", type=["xls", "xlsx"])

if archivo_u221 and archivo_u257 and archivo_sald_rep:

    # Cargar dataframes
    Dic_FAM_REP = cargar_excel(archivo_dic, sheet_name='FAM_REP', dtype={'Famil': str}, parts = True)
    Dic_CTAS_REP = cargar_excel(archivo_dic, sheet_name='CTAS_REP', dtype={'Cuenta short': str, 'Almacén': str, 'Cuenta mayor': str}, parts = True)
    Dic_ALM = cargar_excel(archivo_dic, sheet_name='ALM_REP', dtype={'Almacén': str}, parts = True)

    Stock = cargar_excel(archivo_u221, sheet_name = 'Informe', parts = True, dtype={'Almacé': str})
    Sald_Rep = cargar_excel(archivo_sald_rep, sheet_name = 'Informe', parts = True)
    Compra = cargar_excel(archivo_u257, sheet_name = 'Informe', parts = True, dtype={'T.compr': str}) 

    Dic_FV_REP1 = cargar_excel(archivo_dic, sheet_name='FV_REP_COD', parts = True)
    Dic_FV_REP2 = cargar_excel(archivo_dic, sheet_name='FV_REP_COD2', parts = True)

    Dic_ALM.rename(columns={'Almacén': 'Alm'}, inplace = True)
    Dic_FAM_REP['Familia'] = Dic_FAM_REP['Familia'].replace('NA', 'NR')
    Dic_FAM_REP = Dic_FAM_REP[['Familia', 'GRUPO']]
    Dic_FAM_REP.drop_duplicates(subset=['Familia'], inplace=True)

    Stock['Meses'] = pd.to_numeric(Stock['Meses'], errors='coerce').fillna(0)
    Stock['Valor stock'] = pd.to_numeric(Stock['Valor stock'], errors='coerce').fillna(0)
    Stock.rename(columns={'Famil':"Familia"}, inplace=True)
    Stock['Familia'] = Stock['Familia'].replace('NA', 'NR')
    Stock = pd.merge(Stock, Dic_FAM_REP, on='Familia', how='left')
    Stock['Dias'] = Stock['Meses']*30
    Stock['Rango'] = Stock['Dias'].apply(lambda x: '<90' if x < 90 else '<180' if x < 180 else '<360' if x < 360 else '>360')
    Stock['Fam_Alm'] = Stock['Familia'] + Stock['Almacé']
    Stock = pd.merge(Stock, Dic_CTAS_REP, on='Fam_Alm', how='left')
    Stock = Stock.groupby(['Ciudad', 'Grupo_Contable', 'Rango'], as_index=False)['Valor stock'].sum()
    Stock["Pct_Comp"] = Stock["Valor stock"] / Stock.groupby(['Ciudad', 'Grupo_Contable'])["Valor stock"].transform("sum")
    Stock.drop(columns=['Valor stock'], inplace=True)

    Sald_Rep = Sald_Rep[['Cuenta mayor', 'Saldo']]
    Sald_Rep.rename(columns={'Cuenta mayor': 'Cuenta'}, inplace=True)
    Dic_CTAS_REP.drop_duplicates(subset=['Cuenta mayor'], inplace=True)
    Dic_CTAS_REP.rename(columns={'Cuenta mayor': 'Cuenta'}, inplace=True)
    Dic_CTAS_REP = Dic_CTAS_REP[['Cuenta', 'Ciudad', 'Grupo_Contable']]

    Sald_Rep = pd.merge(Sald_Rep, Dic_CTAS_REP, on='Cuenta', how='left')
    Sald_Rep.dropna(subset=['Grupo_Contable'], inplace=True)
    Sald_Rep = Sald_Rep.groupby(['Ciudad', 'Grupo_Contable'], as_index=False)['Saldo'].sum()

    Stock = pd.merge(Stock, Sald_Rep, on = ['Ciudad', 'Grupo_Contable'], how = 'left')
    Stock['Comp_Saldo'] = Stock['Pct_Comp'] * Stock['Saldo']
    Stock.drop(columns=['Saldo', 'Pct_Comp'], inplace=True)
    Sald_Rep.rename(columns={'Grupo_Contable':'FAMILIA_FV'}, inplace = True)
    
    df_viz = pd.pivot_table(Stock, values='Comp_Saldo', index = ['Ciudad', 'Grupo_Contable'], columns='Rango', aggfunc='sum')
    df_viz = df_viz.fillna(0)
    df_viz = df_viz[['<90', '<180', '<360', '>360']]
    cols_format = ['<90', '<180', '<360', '>360']
    for col in cols_format:
        df_viz[col] = df_viz[col].apply(lambda x: '{:,.2f}'.format(x))
    
    st.write("El análisis de stock presenta el siguiente resultado:")
    st.dataframe(df_viz)

    Compra.rename(columns={'Tot.compra': 'Saldo', 'Almacé': 'Alm'}, inplace = True)
    Compra = Compra[Compra['T.compr'].isin(['1', '2'])]
    Compra['Saldo'] = pd.to_numeric(Compra['Saldo'], errors='coerce').fillna(0)
    Compra = Compra[['Alm', 'Familia', 'Saldo']]
    Compra['Familia'] = Compra['Familia'].replace('NA', 'NR')
    Compra = pd.merge(Compra, Dic_ALM, on='Alm', how='left')
    Compra = pd.merge(Compra, Dic_FAM_REP, on='Familia', how='left')
    Compra = Compra[['Ciudad', 'GRUPO', 'Saldo']]
    Compra.rename(columns={'GRUPO': 'FAMILIA_FV'}, inplace = True)
    Compra = Compra.groupby(['Ciudad', 'FAMILIA_FV'], as_index=False)['Saldo'].sum()

    FV_REP_1 = pd.merge(Dic_FV_REP1, Compra, on=['Ciudad','FAMILIA_FV'], how='inner')
    FV_REP_1 = FV_REP_1[['COD_FV', 'Saldo']]

    FV_REP_2 = pd.merge(Dic_FV_REP2, Stock, on=['Ciudad','Grupo_Contable', 'Rango'], how='left')
    FV_REP_2['Comp_Saldo'] = FV_REP_2['Comp_Saldo'].fillna(0)
    FV_REP_2 = FV_REP_2[['COD_FV', 'Comp_Saldo']]
    FV_REP_2.rename(columns={"Comp_Saldo": "Saldo"}, inplace=True)

    # Reporte Final
    del Compra, Sald_Rep, Stock, Dic_FAM_REP, Dic_ALM, Dic_FV_REP2, Dic_FV_REP1, Dic_CTAS_REP, archivo_u221, archivo_u257, archivo_sald_rep
    
    FV_Parts_Report = pd.concat([FV_REP_1, FV_REP_2], ignore_index=True)
    
    # Botón para descargar CSV
    csv = FV_Parts_Report.to_csv(index=False).encode('utf-8')
    st.download_button(label="Descargar Reporte", data=csv, file_name="FV_Parts_Report.csv", mime='text/csv')


    st.dataframe(FV_Parts_Report)