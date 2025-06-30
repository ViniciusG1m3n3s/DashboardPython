import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from .calculations import calcular_tmo_equipe_cadastro, calcular_ranking_atualizacao, calcular_ranking_distribuicao, calcular_ranking_auditoria, calcular_ranking_cadastro_orgaos,calcular_ranking_cadastro_oficios, calcular_ranking_cadastro_pre,calcular_ranking_cadastro_judicial, gerar_ficha_html_analista, contar_desvios, exibir_cadastro_atualizacao_por_modulo, calcular_cadastro_atualizacao_por_modulo, obter_maior_quantidade_por_fila, exibir_grafico_desvios_auditoria, exibir_melhor_analista_por_fila, exibir_maior_quantidade_por_fila, calcular_e_exibir_tmo_cadastro_atualizacao_por_fila, format_timedelta_hms,exibir_grafico_tmo_analista_por_mes, format_timedelta_grafico_tmo_analista, obter_melhor_analista_por_fila, exibir_grafico_tempo_ocioso_por_dia, calcular_producao_email_detalhada, calcular_producao_agrupada, exportar_planilha_com_tmo_completo, gerar_relatorio_html, download_html, download_html_tmo, gerar_relatorio_html_tmo,  calcular_tmo_equipe_atualizado, calcular_produtividade_diaria, calcular_tmo_por_dia_cadastro, calcular_produtividade_diaria_cadastro, calcular_tmo_por_dia, convert_to_timedelta_for_calculations, convert_to_datetime_for_calculations, save_data, load_data, format_timedelta, calcular_ranking, calcular_filas_analista, calcular_metrica_analista, calcular_carteiras_analista,exportar_relatorio_detalhado_por_analista, get_points_of_attention, calcular_tmo_por_carteira, calcular_tmo, calcular_e_exibir_tmo_por_fila, calcular_tmo_por_mes, exibir_tmo_por_mes, exibir_dataframe_tmo_formatado, export_dataframe, calcular_tempo_ocioso_por_analista, calcular_melhor_tmo_por_dia, calcular_melhor_dia_por_cadastro, exibir_tmo_por_mes_analista, exportar_planilha_com_tmo, calcular_tmo_geral, calcular_tmo_cadastro, calcular_tempo_ocioso, gerar_relatorio_tmo_completo
from .charts import plot_produtividade_diaria, plot_grafico_desvios, plot_tmo_por_dia_cadastro, plot_tmo_por_dia_cadastro, exibir_grafico_tp_causa, plot_produtividade_diaria_cadastros, plot_tmo_por_dia, plot_status_pie, grafico_tmo, grafico_status_analista, exibir_grafico_filas_realizadas, exibir_grafico_tmo_por_dia, exibir_grafico_quantidade_por_dia
from datetime import datetime
import difflib
from Amil.diario import diario

def dashboard():
    hide_footer_style = """ 
    <style>
        ._link_gzau3_10 {
            display: none;
        }
    </style>
    """
    st.markdown(hide_footer_style, unsafe_allow_html=True)
    
    hide_streamlit_style = """
                    <style>
                    div[data-testid="stToolbar"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    div[data-testid="stDecoration"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    div[data-testid="stStatusWidget"] {
                    visibility: hidden;
                    height: 0%;
                    position: fixed;
                    }
                    #MainMenu {
                    visibility: hidden;
                    height: 0%;
                    }
                    header {
                    visibility: hidden;
                    height: 0%;
                    }
                    footer {
                    visibility: hidden;
                    height: 0%;
                    }
                    </style>
                    """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True)
    
    hide_github_icon = """
    <style>
    #GithubIcon {
    visibility: hidden;
    }
    </style>
    """
    st.markdown(hide_github_icon, unsafe_allow_html=True)
    
    background_image_css = """
    <style>
    header {
    background-color: rgba(255, 255, 255, 0); /* Torna o fundo do cabeçalho transparente */
    color: transparent; /* Remove o texto do cabeçalho (opcional) */
    box-shadow: none; /* Remove a sombra (opcional) */
    display: none;
    }
    [data-testid="stHeader"] {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    .stToolbarActions st-emotion-cache-1p1m4ay e3i9eg820 {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    
    .stAppToolbar st-emotion-cache-15ecox0 e10jh26i2 {
        background-color: rgba(255, 255, 255, 0); /* Transparente no novo identificador */
    }
    
        /* Seleciona a barra do Streamlit */
    .st-emotion-cache-15ecox0 {
        background-color: rgba(255, 255, 255, 0.0) !important; /* Transparente */
        box-shadow: none !important; /* Remove sombra */
    }

    /* Remove bordas e sombras adicionais */
    .stToolbarActions, .st-emotion-cache-czk5ss {
        background-color: rgba(255, 255, 255, 0.0) !important;
        box-shadow: none !important;
    }
    
    .css-1jc7ptx, .e1ewe7hr3, .viewerBadge_container__1QSob,
    .styles_viewerBadge__1yB5_, .viewerBadge_link__1S137,
    .viewerBadge_text__1JaDK {
        display: none;
    }
    
    #MainMenu {
        visibility: hidden;
        }   
    
    #GithubIcon {
        visibility: hidden;
    }
    </style>
    """
    st.markdown(background_image_css, unsafe_allow_html=True)
    
    hide_streamlit_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {visibility: hidden;}
                </style>
                """
    st.markdown(hide_streamlit_style, unsafe_allow_html=True) 
    
    # Carregar dados
    usuario_logado = st.session_state.usuario_logado
    df_total = load_data(usuario_logado)

    # Sidebar
    st.sidebar.header("Navegação")
    opcao_selecionada = st.sidebar.selectbox("Escolha uma visão", ["Visão Geral", "Métricas Individuais", "Diário de Bordo"])
    
    # Carregar nova planilha
    uploaded_file = st.sidebar.file_uploader("Carregar nova planilha", type=["xlsx"])

    if uploaded_file is not None:
        df_new = pd.read_excel(uploaded_file)
        df_total = pd.concat([df_total, df_new], ignore_index=True)
        save_data(df_total, usuario_logado)
        st.sidebar.success(f'Arquivo "{uploaded_file.name}" carregado com sucesso!')
        

    if usuario_logado == "andrew@unimed" and not hasattr(st.session_state, 'bianca_welcomed'):
        st.toast("Bem-vindo, Andrew!", icon=":material/account_circle:")
        st.session_state.bianca_welcomed = True

    # Converte para cálculos temporários
    df_total = convert_to_timedelta_for_calculations(df_total)
    df_total = convert_to_datetime_for_calculations(df_total)
    
    ms = st.session_state

    # Verifique se a chave 'themes' existe no session_state
    if "themes" not in ms:
        ms.themes = {
            "current_theme": "light",  # Tema padrão
            "refreshed": True,
            
            # Definições para o tema claro
            "light": {
                "theme.base": "light",  # Tema base claro
                "theme.primaryColor": "#ff521a",  # Cor primária
                "theme.backgroundColor": "#FFFFFF",
                "theme.secondaryBackgroundColor": "#F0F2F6",  # Cor de fundo
                "theme.textColor": "#31333F",  # Cor do texto
                "button_face": ":material/light_mode:",  # Ícone para o botão
                "logo": "logo_light.png",  # Logo para o tema claro
            },
            
            # Definições para o tema escuro
            "dark": {
                "theme.base": "dark",  # Tema base escuro
                "theme.primaryColor": "#ff521a",  # Cor primária
                "theme.backgroundColor": "black",
                "theme.secondaryBackgroundColor": "#262730",  # Cor de fundo
                "theme.textColor": "white",  # Cor do texto
                "button_face": ":material/dark_mode:",  # Ícone para alternar para o tema claro
                "logo": "logo_dark.png",  # Logo para o tema escuro
            }
        }

    # Função para alterar o tema
    def ChangeTheme():
        # Obter o tema anterior
        previous_theme = ms.themes["current_theme"]
            
        # Obter o dicionário de configurações do tema baseado no tema atual
        theme_dict = ms.themes["light"] if ms.themes["current_theme"] == "light" else ms.themes["dark"]
            
        # Definir as opções do tema com base nas configurações
        for key, value in theme_dict.items():
            if key.startswith("theme"):
                st._config.set_option(key, value)
        
        # Alterar o tema atual
        if previous_theme == "dark":
            ms.themes["current_theme"] = "light"
        else:
            ms.themes["current_theme"] = "dark"
            
        ms.themes["refreshed"] = False

    # Definindo o botão para troca de tema
    btn_face = ms.themes["light"]["button_face"] if ms.themes["current_theme"] == "light" else ms.themes["dark"]["button_face"]
    st.sidebar.button(btn_face, on_click=ChangeTheme)

    # Lógica para exibir o logo baseado no tema
    if ms.themes["current_theme"] == "light":
        st.logo("https://finchsolucoes.com.br/img/eb28739f-bef7-4366-9a17-6d629cf5e0d9.png")  # Logo para o tema claro
    else:
        st.logo("https://finchsolucoes.com.br/img/fefdd9df-1bd3-4107-ab22-f06d392c1f55.png")  # Logo para o tema escuro

    # Rerun após a alteração do tema
    if ms.themes["refreshed"] == False:
        ms.themes["refreshed"] = True
        st.rerun()
        
    custom_colors = ['#ff571c', '#7f2b0e', '#4c1908', '#ff884d', '#a34b28', '#331309']
    
    if opcao_selecionada == "Visão Geral":
        
        st.title("Produtividade Geral")

        # Filtros de data
        min_date = df_total['DATA DE CONCLUSÃO DA TAREFA'].min().date() if not df_total.empty else datetime.today().date()
        max_date = df_total['DATA DE CONCLUSÃO DA TAREFA'].max().date() if not df_total.empty else datetime.today().date()
        
        st.subheader("Filtro por Data")
        col1, col2 = st.columns(2)
        with col1:
            data_inicial = st.date_input("Data Inicial", min_date)
        with col2:
            data_final = st.date_input("Data Final", max_date)

        if data_inicial > data_final:
            st.sidebar.error("A data inicial não pode ser posterior à data final!")

        df_total = df_total[(df_total['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) & (df_total['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final)]

        # Métricas de produtividade
        total_finalizados = len(df_total[df_total['FINALIZAÇÃO'] == 'CADASTRADO'])
        total_atualizados = len(df_total[df_total['FINALIZAÇÃO'] == 'ATUALIZADO'])
        total_distribuidos = len(df_total[df_total['FINALIZAÇÃO'] == 'REALIZADO'])
        total_auditado = len(df_total[df_total['FINALIZAÇÃO'] == 'AUDITADO'])
        total_geral = total_finalizados + total_atualizados + total_distribuidos

        # Calcular tempo médio geral, verificando se o total geral é maior que zero
        if total_geral > 0:
            tempo_medio = (
                df_total[df_total['FINALIZAÇÃO'] == 'CADASTRADO']['TEMPO MÉDIO OPERACIONAL'].sum() +
                df_total[df_total['FINALIZAÇÃO'] == 'ATUALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum() +
                df_total[df_total['FINALIZAÇÃO'] == 'REALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum()
            ) / total_geral
        else:
            tempo_medio = pd.Timedelta(0)  # Define como 0 se não houver dados

        # Calcular tempo médio de cadastros, verificando se o total de cadastros é maior que zero
        if total_finalizados > 0:
            tempo_medio_cadastros = (
                df_total[df_total['FINALIZAÇÃO'] == 'CADASTRADO']['TEMPO MÉDIO OPERACIONAL'].sum()
            ) / total_finalizados
        else:
            tempo_medio_cadastros = pd.Timedelta(0)

        # Calcular tempo médio de atualizações, verificando se o total de atualizações é maior que zero
        if total_atualizados > 0:
            tempo_medio_autalizacoes = (
                df_total[df_total['FINALIZAÇÃO'] == 'ATUALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum()
            ) / total_atualizados
        else:
            tempo_medio_autalizacoes = pd.Timedelta(0)

        # Calcular tempo médio de distribuições, verificando se o total de distribuições é maior que zero
        if total_distribuidos > 0:
            tempo_medio_distribuicoes = (
                df_total[df_total['FINALIZAÇÃO'] == 'REALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum()
            ) / total_distribuidos
        else:
            tempo_medio_distribuicoes = pd.Timedelta(0)
            
        if total_auditado > 0:
            tempo_medio_auditoria = (
                df_total[df_total['FINALIZAÇÃO'] == 'AUDITADO']['TEMPO MÉDIO OPERACIONAL'].sum()
            ) / total_auditado
        else:
            tempo_medio_auditoria = pd.Timedelta(0)
            
        st.write(
            """
            <style>
            [data-testid="stMetricDelta"] svg {
                display: none;  
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            with st.container(border=True):
                st.metric("Total Geral", total_geral, delta=f"Tempo Médio - " + format_timedelta(tempo_medio), delta_color="off", help="Engloba todas as tarefas finalizadas e exibe o tempo médio geral.")
        with col2:
            with st.container(border=True):
                st.metric("Total Cadastros", total_finalizados, delta=f"Tempo Médio - " + format_timedelta(tempo_medio_cadastros), delta_color="off", help="Tempo médio das tarefas cadastradas.")
        with col3:
            with st.container(border=True):
                st.metric("Total Atualizações", total_atualizados, delta=f"Tempo Médio - " + format_timedelta(tempo_medio_autalizacoes), delta_color="off", help="Tempo médio das tarefas atualizadas.")
        with col4:
            with st.container(border=True):
                st.metric("Total Auditoria", total_auditado, delta=f"Tempo Médio - " + format_timedelta(tempo_medio_auditoria), delta_color="off", help="Tempo médio das tarefas auditadas.")

        # Expander com Total Geral --- Sendo a soma de todos os cadastros, reclassificados e andamentos
        with st.expander("Tempo Médio por Fila"):
            df_tmo_por_carteira = calcular_tmo_por_carteira(df_total)
            if isinstance(df_tmo_por_carteira, str):
                st.write(df_tmo_por_carteira)  # Exibe mensagem de erro se as colunas não existirem
            else:
                st.dataframe(df_tmo_por_carteira, use_container_width=True, hide_index=True)
        
        # Exibição na Dashboard
        with st.expander("Produção - Resumo por Grupo"):
            df_producao_agrupada = calcular_producao_agrupada(df_total)
            if isinstance(df_producao_agrupada, str):
                st.write(df_producao_agrupada)
            else:
                st.dataframe(df_producao_agrupada, use_container_width=True, hide_index=True)
        
        with st.expander("Cadastros vs Atualizações por Módulo"):
            exibir_cadastro_atualizacao_por_modulo(df_total)
            
        with st.expander("Produção - Detalhamento Grupo E-MAIL"):
            df_producao_email = calcular_producao_email_detalhada(df_total)
            if isinstance(df_producao_email, str):
                st.write(df_producao_email)
            else:
                st.dataframe(df_producao_email, use_container_width=True, hide_index=True)

        # Calculando e exibindo gráficos
        df_produtividade = calcular_produtividade_diaria(df_total)
        
        df_produtividade_cadastro = calcular_produtividade_diaria_cadastro(df_total)
        
        df_tmo = calcular_tmo_por_dia(df_total)  # Certifique-se de que essa função retorne os dados necessários para o gráfico
        
        df_tmo_cadastro = calcular_tmo_por_dia_cadastro(df_total)  # Certifique-se de que essa função retorne os dados necessários para o gráfico
        
        with st.expander("Desvios Auditoria"):
            st.subheader("Desvios Auditados")
            exibir_grafico_desvios_auditoria(df_total)
        
        col1, col2 = st.columns(2)
        
        with col1:
            
            tab1, tab2 = st.tabs(["Produtividade Diária", "Cadastros e Atualizações Diários"])
            
            with tab1:

                with st.container(border=True):
                    st.subheader("Produtividade Diária - Total das Tarefas Finalizadas")
                    fig_produtividade = plot_produtividade_diaria(df_produtividade, custom_colors)
                    if fig_produtividade:
                        st.plotly_chart(fig_produtividade)
            
            with tab2:
                with st.container(border=True):
                    st.subheader("Produtividade Diária - Cadastros e Atualizações")
                    fig_produtividade = plot_produtividade_diaria_cadastros(df_produtividade_cadastro, custom_colors)
                    if fig_produtividade:
                        st.plotly_chart(fig_produtividade)
                        
        with col2:
        
            tab1, tab2 = st.tabs(["TMO Geral Diário", "TMO Cadastro Diário"])
            
            with tab1:
                with st.container(border=True):
                    st.subheader("Tempo Médio Operacional Diario - Geral")
                    fig_tmo = plot_tmo_por_dia(df_tmo, custom_colors)
                    if fig_tmo:
                        st.plotly_chart(fig_tmo)
                        
            with tab2:
                with st.container(border=True):
                    st.subheader("Tempo Médio Operacional Diário - Cadastros")
                    fig_tmo = plot_tmo_por_dia_cadastro(df_tmo_cadastro, custom_colors)
                    if fig_tmo:
                        st.plotly_chart(fig_tmo)
                            
        with st.expander("Tempo Médio Operacional por Mês"):
                st.subheader("Tempo Médio Operacional por Mês")
                exibir_tmo_por_mes(df_total)
                # Exibir o DataFrame formatado na seção correspondente
                
                #Grafico de TMO por Analista
                df_tmo_analista = calcular_tmo(df_total)

        with st.container(border=True):
            # Filtro de analistas
            st.subheader("Tempo Médio Operacional por Analista")
            analistas = df_tmo_analista['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

            # Lista de usuários a serem ignorados
            usuarios_excluidos = [
                "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
            ]

            # Seleciona apenas os analistas válidos
            analistas_filtrados = [
                a for a in analistas if "_ter" not in a and a not in usuarios_excluidos
            ]

            selected_analistas = st.multiselect(
                "Selecione os Analistas:",
                options=analistas,
                default=analistas_filtrados
            )
            # Mostrar o gráfico de TMO
            df_tmo_analista_filtered = df_tmo_analista[df_tmo_analista['USUÁRIO QUE CONCLUIU A TAREFA'].isin(selected_analistas)]
            fig_tmo_analista = grafico_tmo(df_tmo_analista_filtered, custom_colors)
            if fig_tmo_analista:
                st.plotly_chart(fig_tmo_analista)  
            else:   
                st.write("Nenhum analista selecionado")
                
        with st.expander("Melhor Analista por Fila"):
            st.subheader("Melhor Analista por Fila (TMO e Quantidade)")
            col1, col2 = st.columns(2)
            with col1:
                with st.container(border=True):
                    exibir_melhor_analista_por_fila(df_total)
            with col2:
                with st.container(border=True):
                    exibir_maior_quantidade_por_fila(df_total)
        
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(["Ranking Geral", "Ranking Cadastro", "Ranking Atualizações","Ranking Pré-Cadastro", "Ranking Ofícios", "Ranking Demais Órgãos", "Ranking Auditoria", "Ranking Distribuição"])
        
        with tab1:
            with st.container(border=True):  
                st.subheader("Ranking de Geral")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking"
                )

                # Calcular o ranking
                styled_df_ranking = calcular_ranking(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking, width=2000, hide_index=True)
                
                    # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)
        
        with tab2:
            with st.container(border=True):  
                st.subheader("Ranking Cadastro")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_cadastro"
                )

                # Calcular o ranking
                styled_df_ranking_cadastro = calcular_ranking_cadastro_judicial(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_cadastro, width=2000, hide_index=True)
                
            # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)
            
        with tab3:
            with st.container(border=True):  
                st.subheader("Ranking Atualização")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_atualizacao"
                )

                # Calcular o ranking
                styled_df_ranking_atualizado = calcular_ranking_atualizacao(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_atualizado, width=2000, hide_index=True)
                
            # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)
            
        with tab4:
            with st.container(border=True):  
                st.subheader("Ranking Pré-Cadastro")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_cadastro_pre"
                )

                # Calcular o ranking
                styled_df_ranking_pre_cadastro = calcular_ranking_cadastro_pre(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_pre_cadastro, width=2000, hide_index=True)
        
        with tab5:
            with st.container(border=True):  
                st.subheader("Ranking Ofícios")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_cadastro_oficios"
                )

                # Calcular o ranking
                styled_df_ranking_cadastro_oficios = calcular_ranking_cadastro_oficios(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_cadastro_oficios, width=2000, hide_index=True)
                
                    # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)

        with tab6:
            with st.container(border=True):  
                st.subheader("Ranking Ofícios e Demais Órgãos")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "biancabazolli", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_cadastro_orgaos"
                )

                # Calcular o ranking
                styled_df_ranking_cadastro_orgaos = calcular_ranking_cadastro_orgaos(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_cadastro_orgaos, width=2000, hide_index=True)
                
                    # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)

        with tab7:
            with st.container(border=True):  
                st.subheader("Ranking Auditoria")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "ingridvieira_amil", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_auditoria"
                )

                # Calcular o ranking
                styled_df_ranking_auditoria = calcular_ranking_auditoria(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_auditoria, width=2000, hide_index=True)
                
                    # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)

        with tab8:
            with st.container(border=True):  
                st.subheader("Ranking Distribuição")
                
                            # Selecione os usuários
                users = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

                # Lista de usuários a serem excluídos da seleção padrão
                usuarios_excluidos = [
                    "viniciusgimenes_amil", "biancamaia", "bayrabraz",
                    "barbaralopes_amil", "andrewcossi", "beatrizromao_amil", "eloabernardo", "alexandredomingues", "sabrinalira"
                ]

                # Filtrar para seleção padrão (sem _ter e sem os nomes excluídos)
                users_filtrados = [
                    u for u in users if "_ter" not in u and u not in usuarios_excluidos
                ]

                selected_users = st.multiselect(
                    "Selecione os Analistas:",
                    options=users,
                    default=users_filtrados,
                    key="multiselect_ranking_distribuicao"
                )

                # Calcular o ranking
                styled_df_ranking_distribuicao = calcular_ranking_distribuicao(df_total, selected_users)
                
                # Exibir a tabela de ranking
                st.dataframe(styled_df_ranking_distribuicao, width=2000, hide_index=True)
                
                    # Injetando CSS e JavaScript para aumentar o tamanho do modal
            st.markdown("""
                <style>
                    div [tabindex="-1"] {
                        width: 80% !important;
                    }
                </style>
            """, unsafe_allow_html=True)
                        
        # Função para exibir o Power BI no modal
        @st.dialog("BI - Qualidade AMIL", width="large")
        def abrir_bi():

            iframe_code = """
            <iframe title="BI_QUALIDADE_AMIL" width="100%" height="600" 
                    src="https://app.powerbi.com/reportEmbed?reportId=c023e127-5bc0-46e5-8589-2196b78aa72f&autoAuth=true&ctid=0a336fa9-e4fe-4a46-8908-98290f728ce3" 
                    frameborder="0" allowFullScreen="true"></iframe>
            """

            st.components.v1.html(iframe_code, height=600)

        # Botão para abrir o modal
        if st.button("Abrir Power BI"):
            abrir_bi()
        
        with st.expander("Exportar Dados"):
            try:    
                # Seleção do período
                data_inicial_relatorio = st.date_input(
                    "Data Inicial Relatório", 
                    df_total['DATA DE CONCLUSÃO DA TAREFA'].min().date()
                )
                data_final_relatorio = st.date_input(
                    "Data Final Relatório", 
                    df_total['DATA DE CONCLUSÃO DA TAREFA'].max().date()
                )

                # Seleção de analistas
                analistas_disponiveis = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].unique()
                analistas_selecionados = st.multiselect(
                    "Selecione os analistas", 
                    options=analistas_disponiveis, 
                    default=analistas_disponiveis
                )

                # Seleção do tipo de TMO
                tmo_tipo = st.radio(
                    "Selecione o tipo de TMO para exportação:",
                    options=['GERAL', 'CADASTRADO', 'CADASTRADO_DETALHADO'],
                    index=0,
                    help=(
                        "Escolha 'GERAL' para considerar todas as tarefas, "
                        "'CADASTRADO' para considerar apenas finalizações de CADASTRO, "
                        "ou 'CADASTRADO_DETALHADO' para incluir detalhes dos tipos de 'TP CAUSA (TP COMPLEMENTO)'."
                    )
                )

                # Botão para exportar a planilha padrão
                if st.button("Exportar Planilha Resumida"):
                    periodo_selecionado = (data_inicial_relatorio, data_final_relatorio)
                    exportar_planilha_com_tmo(df_total, periodo_selecionado, analistas_selecionados, tmo_tipo)

                # Botão para exportar a planilha detalhada por analista
                if st.button("Exportar Relatório Detalhado por Analista"):
                    periodo_selecionado = (data_inicial_relatorio, data_final_relatorio)
                    exportar_relatorio_detalhado_por_analista(df_total, periodo_selecionado, analistas_selecionados)
                    
                # Adicionar botão de exportação para o novo relatório
                if st.button("Exportar Planilha Completa de TMO"):
                    periodo_selecionado = (data_inicial_relatorio, data_final_relatorio)
                    exportar_planilha_com_tmo_completo(df_total, periodo_selecionado, analistas_selecionados)

            except ValueError as e:
                st.warning("Ocorreu um erro ao processar as datas. Verifique se as informações de data estão corretas no seu arquivo. Detalhes do erro:")
                st.code(str(e))

            except Exception as e:
                st.warning("Ocorreu um erro inesperado. Por favor, tente novamente. Detalhes do erro:")
                st.code(str(e))

        if not df_total.empty:
            with st.expander("Exportar Relatório de TMO em HTML"):
                # 🔹 Seleção de períodos antes e depois da mudança
                st.subheader("Selecione os períodos para comparação")
                col1, col2 = st.columns(2)

                with col1:
                    data_inicio_antes = st.date_input("Data Inicial Antes", df_total['DATA DE CONCLUSÃO DA TAREFA'].min().date())
                    data_fim_antes = st.date_input("Data Final Antes", df_total['DATA DE CONCLUSÃO DA TAREFA'].max().date())

                with col2:
                    data_inicio_depois = st.date_input("Data Inicial Depois", df_total['DATA DE CONCLUSÃO DA TAREFA'].min().date())
                    data_fim_depois = st.date_input("Data Final Depois", df_total['DATA DE CONCLUSÃO DA TAREFA'].max().date())

                # 🔹 Seleção de usuários
                usuarios_disponiveis = df_total['USUÁRIO QUE CONCLUIU A TAREFA'].unique()
                usuarios_selecionados = st.multiselect(
                    "Selecione os usuários para o relatório",
                    options=usuarios_disponiveis,
                    default=usuarios_disponiveis
                )

                # 🔹 Botão para baixar o HTML
                if st.button("Gerar e Baixar Relatório HTML"):
                    download_html(df_total, data_inicio_antes, data_fim_antes, data_inicio_depois, data_fim_depois, usuarios_selecionados)
        
    elif opcao_selecionada == "Métricas Individuais":
        st.title("Métricas Individuais")
        
        # Filtro de data
        st.subheader("Filtro por Data")
        min_date = df_total['DATA DE CONCLUSÃO DA TAREFA'].min().date() if not df_total.empty else datetime.today().date()
        max_date = df_total['DATA DE CONCLUSÃO DA TAREFA'].max().date() if not df_total.empty else datetime.today().date()

        col1, col2 = st.columns(2)
        with col1:
            data_inicial = st.date_input("Data Inicial", min_date)
        with col2:
            data_final = st.date_input("Data Final", max_date)

        if data_inicial > data_final:
            st.error("A data inicial não pode ser posterior à data final!")

        df_total = df_total[(df_total['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) & (df_total['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final)]
        analista_selecionado = st.selectbox('Selecione o analista', df_total['USUÁRIO QUE CONCLUIU A TAREFA'].unique())
        df_analista = df_total[df_total['USUÁRIO QUE CONCLUIU A TAREFA'] == analista_selecionado].copy()

        # Chama as funções de cálculo
        tmo_equipe_cadastro = calcular_tmo_equipe_cadastro(df_total)
        tmo_equipe_atualizacao = calcular_tmo_equipe_atualizado(df_total)
        
        total_finalizados_analista, total_atualizado_analista, tempo_medio_analista, tmo_cadastrado_analista, tmo_atualizado_analista, total_realizados_analista, media_cadastros_por_dia, dias_trabalhados = calcular_metrica_analista(df_analista)

        # Define valores padrão caso as variáveis retornem como None
        if total_finalizados_analista is None:
            total_finalizados_analista = 0
        if total_atualizado_analista is None:
            total_atualizado_analista = 0
        if total_realizados_analista is None:
            total_realizados_analista = 0
            
        st.write(
            """
            <style>
            [data-testid="stMetricDelta"] svg {
                display: none;  
            }
            </style>
            """,
            unsafe_allow_html=True,
        )
            
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            with st.container(border=True):
                st.metric("Total Geral", total_finalizados_analista+total_atualizado_analista+total_realizados_analista, f"Tempo Médio - {format_timedelta(tempo_medio_analista)}", delta_color="off")  
        with col2:
            with st.container(border=True):
                st.metric("Total Cadastrados", total_finalizados_analista, f"Tempo Médio - {format_timedelta(tmo_cadastrado_analista)}",  delta_color="off")
        with col3:
            with st.container(border=True):
                st.metric("Total Atualizado", total_atualizado_analista, f"Tempo Médio - {format_timedelta(tmo_atualizado_analista)}",  delta_color="off")
        with col4:
            with st.container(border=True):
                st.metric("Média de Cadastros", media_cadastros_por_dia, f"Dias Trabalhados - {dias_trabalhados}",  delta_color="off")
        
        if tmo_cadastrado_analista is not None and tmo_equipe_cadastro is not None:
            if tmo_cadastrado_analista > tmo_equipe_cadastro:
                st.toast(f"O TMO de Cadastro de {analista_selecionado} ({format_timedelta(tmo_cadastrado_analista)}) excede o tempo médio da equipe ({format_timedelta(tmo_equipe_cadastro)}).", icon=":material/warning:")
            else:
                pass
        
        if tmo_atualizado_analista is not None and tmo_equipe_cadastro is not None:
            if tmo_atualizado_analista > tmo_equipe_atualizacao:
                st.toast(f"O TMO de Atualização de {analista_selecionado} ({format_timedelta(tmo_atualizado_analista)}) excede o tempo médio da equipe ({format_timedelta(tmo_equipe_atualizacao)}).", icon=":material/warning:")
            else:
                pass     

        melhor_dia_tmo, melhor_tmo = calcular_melhor_tmo_por_dia(df_analista)
        melhor_dia_cadastro, quantidade_cadastro = calcular_melhor_dia_por_cadastro(df_analista)
    
        with st.expander("Melhor TMO e Quantidade de Cadastro"):
            col1, col2 = st.columns(2)
            with col1:
                with st.container(border=True):
                    if melhor_dia_tmo and melhor_tmo:
                        formatted_tmo = format_timedelta(melhor_tmo)
                        st.metric("Melhor TMO", formatted_tmo, f"Dia {melhor_dia_tmo.strftime('%d/%m/%Y')}")
                    else:
                        st.metric("Melhor TMO", "Sem dados")
            with col2:
                with st.container(border=True):
                    if melhor_dia_cadastro:
                            st.metric("Melhor Dia de Cadastros", quantidade_cadastro, f"Dia {melhor_dia_cadastro.strftime('%d/%m/%Y')}")
                    else:
                        st.metric("Melhor Dia de Cadastros", "Sem dados")
            
        with st.expander("TMO por Fila - Cadastro e Atualização"):
            calcular_e_exibir_tmo_cadastro_atualizacao_por_fila(df_analista, format_timedelta_hms, st)
            
        with st.expander("Tempo Ocioso"):
                st.subheader(f"Tempo Ocioso")
                exibir_grafico_tempo_ocioso_por_dia(df_analista, analista_selecionado, custom_colors, st)
                df_tempo_ocioso = calcular_tempo_ocioso_por_analista(df_analista)
                st.dataframe(df_tempo_ocioso, hide_index=True, use_container_width=True)
                df_tempo_ocioso["Tempo Ocioso Formatado"] = pd.to_timedelta(df_tempo_ocioso["Tempo Ocioso Formatado"])
                # total de segundos do timedelta
                media_tempo_ocioso = df_tempo_ocioso["Tempo Ocioso Formatado"].mean()
                total_segundos = int(media_tempo_ocioso.total_seconds())

                # converte para horas, minutos e segundos
                horas = total_segundos // 3600
                minutos = (total_segundos % 3600) // 60
                segundos = total_segundos % 60

                # formata como 00:06:34
                tempo_formatado = f"{horas:02}:{minutos:02}:{segundos:02}"

                with st.container(border=True):
                    st.metric("Média de Tempo Ocioso", tempo_formatado)
                    
        with st.expander("Evolução TMO"):
            st.subheader(f"Tempo Médio Operacional Mensal")
            exibir_grafico_tmo_analista_por_mes(df_analista, analista_selecionado)
        
        col1, col2 = st.columns(2)
        with col1:
            # Gráfico de TMO por dia usando a função do `graph.py`
            with st.container(border=True):
                st.subheader(f"Tempo Médio Operacional por Dia")
                exibir_grafico_tmo_por_dia(
                df_analista=df_analista,
                analista_selecionado=analista_selecionado,
                calcular_tmo_por_dia=calcular_tmo_por_dia,
                custom_colors=custom_colors,
                st=st
            )

        with col2:
            # Gráfico de TMO por dia usando a função do `graph.py`
            with st.container(border=True):
                st.subheader(f"Quantidade de Tarefas por Dia")
                exibir_grafico_quantidade_por_dia(
                    df_analista=df_analista,
                    analista_selecionado=analista_selecionado,
                    custom_colors=custom_colors,
                    st=st
            )
        
        col1, col2 = st.columns(2)
        with col1:
            with st.container(border=True):
                st.subheader(f"Filas Realizadas")                    
                exibir_grafico_filas_realizadas(
                    df_analista=df_analista,
                    analista_selecionado=analista_selecionado,
                    custom_colors=custom_colors,
                    st=st
                )
        with col2:
            with st.container(border=True):
                st.subheader(f"Tarefas Cadastradas por TP CAUSA")
                exibir_grafico_tp_causa(df_analista, analista_selecionado, custom_colors, st)
        
        def exportar_ficha_html_analista(df, nome_analista, data_inicio, data_fim):
            html = gerar_ficha_html_analista(df, nome_analista, data_inicio, data_fim)
            buffer = BytesIO()
            buffer.write(html.encode("utf-8"))
            buffer.seek(0)

            st.download_button(
                label="📄 Baixar Ficha de Desempenho (HTML)",
                data=buffer,
                file_name=f"ficha_desempenho_{nome_analista.replace(' ', '_')}.html",
                mime="text/html"
            )
        
        exportar_ficha_html_analista(df_analista, analista_selecionado, min_date, max_date)

    elif opcao_selecionada == "Diário de Bordo":
        
        st.header("provisório")

        def responder_dados(pergunta, df):
            pergunta = pergunta.lower()

            # Lista de nomes de analistas conhecidos no DataFrame
            analistas = df['USUÁRIO QUE CONCLUIU A TAREFA'].dropna().unique()

            # Verificar se a pergunta fala de "tmo"
            if "tmo" in pergunta or "tempo médio" in pergunta:
                for nome in analistas:
                    if nome.lower() in pergunta:
                        tmo = df[df['USUÁRIO QUE CONCLUIU A TAREFA'].str.lower() == nome.lower()]['TEMPO MÉDIO OPERACIONAL'].mean()
                        return f"O TMO médio da {nome} é {tmo:.2f}."

                # Caso não encontre o nome, mas peça tmo geral
                if "geral" in pergunta or "todos" in pergunta:
                    tmo = df['TEMPO MÉDIO OPERACIONAL'].mean()
                    return f"O TMO médio geral é {tmo:.2f}."

            # Pergunta sobre produtividade
            elif "mais produtivo" in pergunta or "quem fez mais" in pergunta:
                top = df['USUÁRIO QUE CONCLUIU A TAREFA'].value_counts().idxmax()
                return f"O analista mais produtivo foi: {top}."

            # Total de tarefas
            elif "quantas tarefas" in pergunta or "total de tarefas" in pergunta:
                total = len(df)
                return f"O total de tarefas realizadas foi: {total}."

            return "Desculpe, ainda não sei responder essa pergunta. Tente reformular ou peça ajuda a um analista."
        
        pergunta_usuario = st.chat_input("Envie sua dúvida sobre os dados...")

        if pergunta_usuario:
            st.chat_message("user").write(pergunta_usuario)

            resposta = responder_dados(pergunta_usuario, df_total)
            st.chat_message("assistant").write(resposta)
            
    if st.sidebar.button("Logout", icon=":material/logout:"):
        st.session_state.logado = False
        st.session_state.usuario_logado = None
        st.sidebar.success("Desconectado com sucesso!")
        st.rerun()  # Volta para a tela de login

if __name__ == "__main__":
    dashboard()
