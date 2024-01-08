import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Gerador base consolidada Master")
st.subheader("Caldic LATAM - FP&A")

uploaded_file = st.file_uploader("Inclua o arquivo Excel com os dados da Master", type=['xlsx'])

nome_aba = [
    "LATAM", "LATAM Managerial", 
    "LAS",  "Brazil", "Goaltech", "Corporate LAS", 
    "Corp LAS Brazil", "Corp LAS China", "Argentina", "Chile", "LAN", "Corporate LAN", 
    "Corp LAN Bogota", "Quimicos Basicos", "Corp LAN CSC", "Corp LAN Houston", "Corp LAN China",
    "PCM", "TPC", "Mexico", "CENAM", "Cluster CENAM", "Guatemala", "Honduras", "El Salvador", 
    "Nicaragua", "Costa Rica", "Panama", "ANDEAN", "Cluster ANDEAN", "Colombia", "Peru", "Ecuador", 
    "Corporate LATAM", "Corporate SP", "Corporate Holding", "Corporate Brazil", "GTM Espanha", "TMLA", 
    "Sotro", "AJ", "Corporate Houston", "GTMI-CP", "M&A", "Active", "Bring"
]

def tratar_master(uploaded_file, sheet_name, range_cols, origem):

    df = pd.read_excel(
        uploaded_file,
        sheet_name,
        header = 1,
        skiprows = lambda x: x in [1,2,3,4,6,7,8,9,10,11],
        nrows = 133,
        usecols = range_cols
    )

    df.dropna(axis ='index', how='all', inplace = True)

    df['type'] = origem

    df = df.rename(columns={"USD 000'":'indicador'})

    df  = pd.melt(
        df,
        id_vars=['indicador', 'type'],
        value_name='usd_000',
        var_name='month'
    )

    return df

if uploaded_file and st.button('Consolidar arquivos'):
    with st.spinner('Consolidando arquivos...'):
        # ACTUAL
        dataframes_actual = {}
        range_cols = "B,V:AG" 
        for aba in nome_aba:
            df = tratar_master(uploaded_file, aba, range_cols, 'Actual')
            df['nome_aba'] = aba
            dataframes_actual[aba] = df
        actual = pd.concat(dataframes_actual.values(), ignore_index=True)

        # FORECAST
        dataframes_forecast = {}
        range_cols = "B,AK:AV" 
        for aba in nome_aba:
            df = tratar_master(uploaded_file, aba, range_cols, 'Forecast')
            df['nome_aba'] = aba
            dataframes_forecast[aba] = df
        forecast = pd.concat(dataframes_forecast.values(), ignore_index=True)

        # BUDGET
        dataframes_budget = {}
        range_cols = "B,AZ:BK" 
        for aba in nome_aba:
            df = tratar_master(uploaded_file, aba, range_cols, 'Budget')
            df['nome_aba'] = aba
            dataframes_budget[aba] = df
        budget = pd.concat(dataframes_budget.values(), ignore_index=True)

        # ACTUAL 22
        dataframes_actual22 = {}
        range_cols = "B,BO:BZ" 
        for aba in nome_aba:
            df = tratar_master(uploaded_file, aba, range_cols, 'Actual 2022')
            df['nome_aba'] = aba
            dataframes_actual22[aba] = df
        actual22 = pd.concat(dataframes_actual22.values(), ignore_index=True)

        # CONSOLIDAÇÃO 
        df = pd.concat([actual, forecast, budget, actual22], axis=0)
        df['month'] = pd.to_datetime(df['month']).dt.strftime('%d-%m-%Y')

        st.dataframe(df)

        # EXPORT EXCEL
        df.to_excel('dados_master_consolidado.xlsx', index=False)

        towrite = BytesIO()
        df.to_excel(towrite, index=False)
        towrite.seek(0) 

        st.download_button(label="📥 Download Final Excel",
                        data=towrite,
                        file_name='dados_master_consolidado.xlsx',
                        mime="application/vnd.ms-excel")