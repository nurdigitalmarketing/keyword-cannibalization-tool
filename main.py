import datetime
import io
import streamlit as st
import pandas as pd
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Configurazione della pagina con titolo e icona
st.set_page_config(page_title="Keyword Cannibalization Tool • NUR® Digital Marketing", page_icon="./Nur-simbolo-1080x1080.png")
st.image("./logo_nur_vettoriale.svg", width=100)

st.markdown("# Keyword Cannibalization Tool")
st.markdown("""
Trova le pagine che competono tra loro per la stessa parola chiave utilizzando i dati GSC.
I dati devono contenere informazioni multidimensionali su query (parole chiave) e pagine.
""")

perc_slider = st.slider('Imposta Soglia (es: 80%)', 0, 100, value=80, step=10)

with st.expander("Come funziona?"):
    st.markdown("""
1. Autenticati con Google Search Console
2. Seleziona un dominio e i parametri di ricerca
3. I dati verranno automaticamente estratti e processati per l'analisi della cannibalizzazione
""")

# Funzione di autenticazione Google
def google_auth():
    client_config = {
        "installed": {
            "client_id": st.secrets["installed"]["client_id"],
            "client_secret": st.secrets["installed"]["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://accounts.google.com/o/oauth2/token",
            "redirect_uris": st.secrets["installed"]["redirect_uris"]
        }
    }
    scopes = ["https://www.googleapis.com/auth/webmasters.readonly"]
    flow = Flow.from_client_config(client_config, scopes=scopes, redirect_uri=st.secrets["installed"]["redirect_uris"][0])
    auth_url, _ = flow.authorization_url(prompt="consent")
    return flow, auth_url

# Verifica se l'utente è già autenticato tramite session state
if 'credentials' not in st.session_state:
    # Autenticazione con Google Search Console
    flow, auth_url = google_auth()
    st.markdown(f"[Accedi a Google Search Console]({auth_url})")

    # Usare st.query_params invece di st.experimental_get_query_params
    query_params = st.query_params
    auth_code = query_params.get("code")

    if auth_code:
        try:
            flow.fetch_token(code=auth_code)
            st.session_state['credentials'] = flow.credentials  # Salva le credenziali in sessione
            st.experimental_rerun()  # Ricarica la pagina dopo l'autenticazione
        except Exception as e:
            st.error(f"Errore durante l'autenticazione: {e}")
else:
    # Se l'utente è autenticato, continua con l'estrazione dei dati
    credentials = st.session_state['credentials']
    service = build('webmasters', 'v3', credentials=credentials)

    @st.cache_data
    def list_properties(_service):
        try:
            site_list = _service.sites().list().execute()
            return [site['siteUrl'] for site in site_list.get('siteEntry', [])]
        except Exception as e:
            st.error(f"Errore nel recuperare le proprietà: {e}")
            return []

    properties = list_properties(service)
    if properties:
        selected_property = st.selectbox("Seleziona il dominio", properties)

        # Parametri di ricerca GSC
        search_type = st.selectbox("Tipo di ricerca", ["web", "image", "video"])
        start_date = st.date_input("Data inizio", datetime.date.today() - datetime.timedelta(days=7))
        end_date = st.date_input("Data fine", datetime.date.today())

        @st.cache_data
        def fetch_gsc_data(_service, selected_property, request):
            try:
                response = _service.searchanalytics().query(siteUrl=selected_property, body=request).execute()
                rows = response.get('rows', [])
                # Estrarre i dati e trasformarli in DataFrame
                data = pd.DataFrame(rows)

                # Espandere le chiavi (query e page)
                if not data.empty:
                    data[['query', 'page']] = pd.DataFrame(data['keys'].tolist(), index=data.index)
                    data.drop(columns=['keys'], inplace=True)  # Rimuove la colonna keys
                return data
            except Exception as e:
                st.error(f"Errore nel recuperare i dati da Google Search Console: {e}")
                return pd.DataFrame()

        if st.button("Estrai dati"):
            request = {
                'startDate': start_date.strftime('%Y-%m-%d'),
                'endDate': end_date.strftime('%Y-%m-%d'),
                'dimensions': ['query', 'page'],
                'searchType': search_type
            }
            data = fetch_gsc_data(service, selected_property, request)

            # Selettore delle dimensioni da visualizzare
            if not data.empty:
                columns = data.columns.tolist()
                selected_columns = st.multiselect("Seleziona le colonne da visualizzare", columns, default=columns)

                # Visualizza solo le colonne selezionate
                st.dataframe(data[selected_columns])

                # Scaricare i dati in Excel
                def to_excel(df):
                    output = io.BytesIO()
                    writer = pd.ExcelWriter(output, engine='xlsxwriter')
                    df.to_excel(writer, sheet_name='Cannibalization Data', index=False)
                    writer.close()  # Usa writer.close() invece di save()
                    output.seek(0)
                    return output.getvalue()

                st.download_button('Scarica Analisi', data=to_excel(data[selected_columns]), file_name='cannibalization_data.xlsx')

# Footer
st.markdown("---")
st.markdown("© 2024 [NUR® Digital Marketing](https://www.nur.it/). Tutti i diritti riservati.")
