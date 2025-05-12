import streamlit as st
from login import login
import importlib

st.set_page_config(
    page_title="Dashboard",  # Título da aba do navegador
    page_icon="🟠",  # Favicon, você pode usar um emoji ou um caminho para um arquivo .ico
    layout="wide",  # Layout da página, pode ser "wide" ou "centered"
)

# Verifica se o usuário está logado
if 'logado' not in st.session_state:
    st.session_state.logado = False

# Se não estiver logado, mostra o formulário de login
if not st.session_state.logado:
    if login():
        st.rerun()  # Reinicia a aplicação para carregar a dashboard
else:
    # Se estiver logado, carrega a dashboard específica
    usuario_logado = st.session_state.usuario_logado
    dominio = usuario_logado.split('@')[1]  # Obtém o domínio do e-mail (por exemplo, 'itau' ou 'bv')

    if dominio == "itau":
        # Carrega o dashboard da pasta Itau
        dashboard_itau = importlib.import_module("Itau.dashboard")
        dashboard_itau.dashboard()  # Chama a função 'dashboard' dentro do arquivo dashboard.py da pasta Itau
    elif dominio == "bv":
        # Carrega o dashboard da pasta BV
        dashboard_bv = importlib.import_module("BV.dashboard")
        dashboard_bv.dashboard()  # Chama a função 'dashboard' dentro do arquivo dashboard.py da pasta BV
    elif dominio == "maestro":
        # Carrega o dashboard da pasta BV
        dashboard_bv = importlib.import_module("Maestro.dashboard")
        dashboard_bv.dashboard()  # Chama a função 'dashboard' dentro do arquivo dashboard.py da pasta BV
    elif dominio == "oficios":
        dashboard_oficios = importlib.import_module("Oficios.dashboard")    
        dashboard_oficios.dashboard()
    elif dominio == "amil":
        dashboard_oficios = importlib.import_module("Amil.dashboard")    
        dashboard_oficios.dashboard()
    elif dominio == "unimed":
        dashboard_oficios = importlib.import_module("Unimed.dashboard")    
        dashboard_oficios.dashboard()
    else:
        st.error("Usuário não autorizado para visualizar o dashboard.")
