import pandas as pd
import streamlit as st
from io import BytesIO

def tratar_master(df, origem):
    df = df.dropna(axis ='index', how='all')
    df['Type'] = origem
    df = df.rename(columns={"USD 000'":'Line'})
    df  = pd.melt(
        df,
        id_vars=['Line', 'Type'],
        value_name='usd_000',
        var_name='Month'
    )
    return df

# inicialização do estado
if 'file_processed' not in st.session_state:
    st.session_state['file_processed'] = False

st.title("Gerador base consolidada Master")
st.subheader("Caldic LATAM - FP&A")
with st.expander("Como usar:"):
    st.write("""
             Carregue o arquivo Excel da Master no campo indicado.

             Clique em "Carregar arquivo" e aguarde a execução.
             
             Ao fim do processamento do arquivo, clique em "Consolidar arquivo" para gerar a base consolidada da Master
            """)

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


if uploaded_file:
    if st.button('Carregar arquivos'):
        try:
            range_cols_actual = "B,V:AG"
            sheets_actual = pd.read_excel(
                    uploaded_file,
                    sheet_name=nome_aba,
                    header = 1,
                    skiprows = lambda x: x in [1,2,3,4,6,7,8,9,10],
                    nrows = 126,
                    usecols = range_cols_actual
            )
            range_cols_forecast = "B,AK:AV" 
            sheets_forecast = pd.read_excel(
                    uploaded_file,
                    sheet_name=nome_aba,
                    header = 1,
                    skiprows = lambda x: x in [1,2,3,4,6,7,8,9,10],
                    nrows = 126,
                    usecols = range_cols_forecast
            )
            range_cols_budget = "B,AZ:BK"
            sheets_budget = pd.read_excel(
                    uploaded_file,
                    sheet_name=nome_aba,
                    header = 1,
                    skiprows = lambda x: x in [1,2,3,4,6,7,8,9,10],
                    nrows = 126,
                    usecols = range_cols_budget
            )

            range_cols_actual22 = "B,BO:BZ" 
            sheets_actual22 = pd.read_excel(
                    uploaded_file,
                    sheet_name=nome_aba,
                    header = 1,
                    skiprows = lambda x: x in [1,2,3,4,6,7,8,9,10],
                    nrows = 126,
                    usecols = range_cols_actual22
            )

            st.session_state['sheets_actual'] = sheets_actual
            st.session_state['sheets_forecast'] = sheets_forecast
            st.session_state['sheets_budget'] = sheets_budget
            st.session_state['sheets_actual22'] = sheets_actual22

            st.success("Carregado!")
            st.session_state['file_processed'] = True

        except Exception as e:
            str.error(f"Erro ao processar o arquivo: {e}")


if st.session_state.get('file_processed'):
    if st.button('Consolidar arquivo'):
        with st.spinner('Consolidando arquivos...'):
            try:

                sheets_actual = st.session_state.get('sheets_actual')
                sheets_forecast = st.session_state.get('sheets_forecast')
                sheets_budget = st.session_state.get('sheets_budget')
                sheets_actual22 = st.session_state.get('sheets_actual22')

                # Check if the sheets are available
                if sheets_actual is None or sheets_forecast is None or \
                   sheets_budget is None or sheets_actual22 is None:
                    st.error("Erro: dados não carregados corretamente.")
                
                # ACTUAL
                dataframes_actual = {}
                for aba in nome_aba:
                    df = sheets_actual[aba]
                    df = tratar_master(df, 'Actual 2023')
                    df['nome_aba'] = aba
                    dataframes_actual[aba] = df
                actual = pd.concat(dataframes_actual.values(), ignore_index=True)

                # FORECAST
                dataframes_forecast = {}
                for aba in nome_aba:
                    df = sheets_forecast[aba]
                    df = tratar_master(df, 'Forecast')
                    df['nome_aba'] = aba
                    dataframes_forecast[aba] = df
                forecast = pd.concat(dataframes_forecast.values(), ignore_index=True)

                # BUDGET
                dataframes_budget = {}
                for aba in nome_aba:
                    df = sheets_budget[aba]
                    df = tratar_master(df, 'Budget')
                    df['nome_aba'] = aba
                    dataframes_budget[aba] = df
                budget = pd.concat(dataframes_budget.values(), ignore_index=True)

                # ACTUAL 22
                dataframes_actual22 = {}
                for aba in nome_aba:
                    df = sheets_actual22[aba]
                    df = tratar_master(df, 'Actual 2022')
                    df['nome_aba'] = aba
                    dataframes_actual22[aba] = df
                actual22 = pd.concat(dataframes_actual22.values(), ignore_index=True)

                # CONSOLIDAÇÃO 
                df = pd.concat([actual, forecast, budget, actual22], axis=0)
                df['Month'] = pd.to_datetime(df['Month']).dt.strftime('%d-%m-%Y')
                df = df.loc[df.usd_000 != '-'].copy()
                df = df.loc[~((df.Line == 'Non-recurring') & (df.usd_000 > 0))].copy()

                # EXPORT EXCEL
                towrite = BytesIO()
                df.to_excel(towrite, index=False)
                towrite.seek(0)

                st.dataframe(df)

                st.success("Arquivos consolidados com sucesso!")

                st.download_button(label="📥 Download Excel Consolidado",
                        data=towrite,
                        file_name='dados_master_consolidado.xlsx',
                        mime="application/vnd.ms-excel")
        
                st.session_state['file_processed'] = False
            
            except Exception as e:
                st.error(f"Erro na consolidação: {e}")









