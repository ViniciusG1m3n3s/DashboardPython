import plotly.express as px
import os
import pandas as pd
import streamlit as st
import plotly.graph_objs as go
import streamlit as st

def plot_produtividade_diaria(df_produtividade, custom_colors):
    if df_produtividade.empty or 'Dia' not in df_produtividade.columns or 'Produtividade' not in df_produtividade.columns:
        st.warning("Não há dados para exibir no gráfico de produtividade diária.")
        return None

    # Ordenar e definir período mínimo e máximo
    df_produtividade = df_produtividade.sort_values(by='Dia')
    data_minima = df_produtividade['Dia'].min()
    data_maxima = df_produtividade['Dia'].max()

    # Criar slider interativo para seleção de período
    periodo_selecionado_plot_produtividade = st.slider(
        "Selecione o período",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=21), data_maxima),
        format="DD MMM YYYY"
    )

    # Filtrar os dados pelo período selecionado
    df_filtrado = df_produtividade[
        (df_produtividade['Dia'] >= periodo_selecionado_plot_produtividade[0]) &
        (df_produtividade['Dia'] <= periodo_selecionado_plot_produtividade[1])
    ]

    # Criar gráfico de linha com pontos
    fig = px.line(
        df_filtrado,
        x='Dia',
        y='Produtividade',
        color_discrete_sequence=custom_colors,
        labels={'Produtividade': 'Total de Cadastros', 'Dia': 'Data'},
        line_shape='linear',
        markers=True
    )

    # Melhorar a formatação do hover e eixo X
    fig.update_traces(
        hovertemplate='Data = %{x|%d/%m/%Y}<br>Produtividade = %{y}'
    )

    fig.update_layout(
        xaxis=dict(
            tickvals=df_filtrado['Dia'],
            ticktext=[f"{dia.day}/{dia.month}/{dia.year}" for dia in df_filtrado['Dia']],
            title='Data'
        ),
        yaxis=dict(
            title='Total de Cadastros'
        )
    )

    # Exibir o gráfico na dashboard
    st.plotly_chart(fig, use_container_width=True)
    
def plot_produtividade_diaria_cadastros(df_produtividade_cadastro, custom_colors):
    if df_produtividade_cadastro.empty or 'Dia' not in df_produtividade_cadastro.columns or 'Produtividade' not in df_produtividade_cadastro.columns:
        st.warning("Não há dados para exibir no gráfico de produtividade diária.")
        return None

    # Ordenar e definir período mínimo e máximo
    df_produtividade_cadastro = df_produtividade_cadastro.sort_values(by='Dia')
    data_minima = df_produtividade_cadastro['Dia'].min()
    data_maxima = df_produtividade_cadastro['Dia'].max()

    # Criar slider interativo para seleção de período
    periodo_selecionado_produtividade = st.slider(
        "Selecione o período",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=30), data_maxima),
        format="DD MMM YYYY"
    )

    # Filtrar os dados pelo período selecionado
    df_filtrado = df_produtividade_cadastro[
        (df_produtividade_cadastro['Dia'] >= periodo_selecionado_produtividade[0]) &
        (df_produtividade_cadastro['Dia'] <= periodo_selecionado_produtividade[1])
    ]

    # Criar gráfico de linha com pontos
    fig = px.line(
        df_filtrado,
        x='Dia',
        y='Produtividade',
        color_discrete_sequence=custom_colors,
        labels={'Produtividade': 'Total de Cadastros', 'Dia': 'Data'},
        line_shape='linear',
        markers=True
    )

    # Melhorar a formatação do hover e eixo X
    fig.update_traces(
        hovertemplate='Data = %{x|%d/%m/%Y}<br>Produtividade = %{y}'
    )

    fig.update_layout(
        xaxis=dict(
            tickvals=df_filtrado['Dia'],
            ticktext=[f"{dia.day}/{dia.month}/{dia.year}" for dia in df_filtrado['Dia']],
            title='Data'
        ),
        yaxis=dict(
            title='Total de Cadastros'
        )
    )

    # Exibir o gráfico na dashboard
    st.plotly_chart(fig, use_container_width=True)

import plotly.express as px
import pandas as pd
import streamlit as st

def plot_tmo_por_dia(df_tmo, custom_colors):
    if df_tmo.empty or 'Dia' not in df_tmo.columns or 'TMO' not in df_tmo.columns:
        st.warning("Não há dados para exibir no gráfico de TMO por dia.")
        return None

    # Garantir que não existam NaN ou valores inválidos na coluna 'TMO'
    df_tmo = df_tmo.dropna(subset=['TMO'])

    # Converter TMO para formato HH:MM:SS
    df_tmo['TMO_Formatado'] = df_tmo['TMO'].apply(
        lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}" 
        if pd.notnull(x) else "00:00:00"
    )

    # Definir período mínimo e máximo para o slider
    df_tmo = df_tmo.sort_values(by='Dia')
    data_minima = df_tmo['Dia'].min()
    data_maxima = df_tmo['Dia'].max()

    # Criar slider interativo (nome único para evitar conflito)
    periodo_tmo_selecionado = st.slider(
        "Selecione o período para TMO por Dia",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=30), data_maxima),
        format="DD MMM YYYY"
    )

    # Filtrar os dados pelo período selecionado
    df_tmo_filtrado = df_tmo[
        (df_tmo['Dia'] >= periodo_tmo_selecionado[0]) &
        (df_tmo['Dia'] <= periodo_tmo_selecionado[1])
    ]

    # Criar gráfico de linha com pontos
    fig_tmo_linha = px.line(
        df_tmo_filtrado,
        x='Dia',
        y=df_tmo_filtrado['TMO'].dt.total_seconds() / 60,  # Converter TMO para minutos
        labels={'y': 'Tempo Médio Operacional (min)', 'Dia': 'Data'},
        color_discrete_sequence=custom_colors,
        line_shape='linear',
        markers=True
    )

    # Melhorar a formatação do hover e eixo X
    fig_tmo_linha.update_traces(
        text=df_tmo_filtrado['TMO_Formatado'],
        textposition='top center',
        hovertemplate='Data = %{x|%d/%m/%Y}<br>TMO = %{text}'
    )

    fig_tmo_linha.update_layout(
        xaxis=dict(
            tickvals=df_tmo_filtrado['Dia'],
            ticktext=[f"{dia.day}/{dia.month}/{dia.year}" for dia in df_tmo_filtrado['Dia']],
            title='Data'
        ),
        yaxis=dict(
            title='Tempo Médio Operacional (HH:MM:SS)'
        )
    )

    return fig_tmo_linha


def plot_tmo_por_dia_cadastro(df_tmo_cadastro, custom_colors):
    if df_tmo_cadastro.empty or 'Dia' not in df_tmo_cadastro.columns or 'TMO' not in df_tmo_cadastro.columns:
        st.warning("Não há dados para exibir no gráfico de TMO por dia.")
        return None

    # Garantir que a coluna TMO seja timedelta
    if isinstance(df_tmo_cadastro['TMO'].iloc[0], str):
        df_tmo_cadastro['TMO'] = pd.to_timedelta(df_tmo_cadastro['TMO'])

    # Converter TMO para formato HH:MM:SS
    df_tmo_cadastro['TMO_Formatado'] = df_tmo_cadastro['TMO'].apply(
        lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}"
        if pd.notnull(x) else "00:00:00"
    )

    # Definir período mínimo e máximo para o slider
    df_tmo_cadastro = df_tmo_cadastro.sort_values(by='Dia')
    data_minima = df_tmo_cadastro['Dia'].min()
    data_maxima = df_tmo_cadastro['Dia'].max()

    # Criar slider interativo (nome único para evitar conflito)
    periodo_tmo_cadastro_selecionado = st.slider(
        "Selecione o período para TMO de Cadastro",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=30), data_maxima),
        format="DD MMM YYYY"
    )

    # Filtrar os dados pelo período selecionado
    df_tmo_cadastro_filtrado = df_tmo_cadastro[
        (df_tmo_cadastro['Dia'] >= periodo_tmo_cadastro_selecionado[0]) &
        (df_tmo_cadastro['Dia'] <= periodo_tmo_cadastro_selecionado[1])
    ]

    # Criar gráfico de linha com pontos
    fig_tmo_cadastro_linha = px.line(
        df_tmo_cadastro_filtrado,
        x='Dia',
        y=df_tmo_cadastro_filtrado['TMO'].dt.total_seconds() / 60,  # Converte TMO para minutos
        labels={'y': 'Tempo Médio Operacional (min)', 'Dia': 'Data'},
        color_discrete_sequence=custom_colors,
        line_shape='linear',
        markers=True
    )

    # Melhorar a formatação do hover e eixo X
    fig_tmo_cadastro_linha.update_traces(
        text=df_tmo_cadastro_filtrado['TMO_Formatado'],
        textposition='top center',
        hovertemplate='Data = %{x|%d/%m/%Y}<br>TMO = %{text}'
    )

    fig_tmo_cadastro_linha.update_layout(
        xaxis=dict(
            tickvals=df_tmo_cadastro_filtrado['Dia'],
            ticktext=[f"{dia.day}/{dia.month}/{dia.year}" for dia in df_tmo_cadastro_filtrado['Dia']],
            title='Data'
        ),
        yaxis=dict(
            title='Tempo Médio Operacional (HH:MM:SS)'
        )
    )

    return fig_tmo_cadastro_linha

def plot_status_pie(total_parcial, total_nao_tratada, total_completa, custom_colors):
    fig_status = px.pie(
        names=['Subsídio Parcial', 'Fora do Escopo', 'Subsídio Completo'],
        values=[total_parcial, total_nao_tratada, total_completa],
        color_discrete_sequence=custom_colors
    )
    fig_status.update_traces(
        hovertemplate='Tarefas %{label} = %{value}<extra></extra>',
    )
    fig_status.update_layout(
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5
        )
    )
    return fig_status

def format_timedelta_grafico_tmo(td):
    if pd.isnull(td):
        return "00:00:00"
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

def grafico_tmo(df_tmo_analista, custom_colors):
    # Verifica se o DataFrame está vazio
    if df_tmo_analista.empty:
        return None  # Retorna None se não houver dados
    
    # Certifica-se de que a coluna 'TMO' é do tipo timedelta
    if 'TMO' not in df_tmo_analista or not pd.api.types.is_timedelta64_dtype(df_tmo_analista['TMO']):
        raise ValueError("A coluna 'TMO' precisa estar no formato timedelta. Verifique os dados.")

    # Certifique-se de que 'TMO_Formatado' existe para exibição no gráfico
    if 'TMO_Formatado' not in df_tmo_analista:
        df_tmo_analista['TMO_Formatado'] = df_tmo_analista['TMO'].apply(format_timedelta_grafico_tmo)

    # Cria o gráfico de barras
    fig_tmo_analista = px.bar(
        df_tmo_analista,
        x='USUÁRIO QUE CONCLUIU A TAREFA',
        y=df_tmo_analista['TMO'].dt.total_seconds() / 60,  # TMO em minutos
        labels={
            'y': 'Tempo Médio Operacional (min)', 
            'x': 'Analista', 
            'USUÁRIO QUE CONCLUIU A TAREFA': 'Analista'
        },
        text=df_tmo_analista['TMO_Formatado'],
        color_discrete_sequence=custom_colors or px.colors.qualitative.Set2
    )
    fig_tmo_analista.update_traces(
        textposition='outside',  # Exibe o tempo formatado fora das barras
        hovertemplate='Analista = %{x}<br>TMO = %{text}<extra></extra>',
        text=df_tmo_analista['TMO_Formatado']
    )
    return fig_tmo_analista

def grafico_status_analista(total_parcial_analista, total_fora_analista, total_completo_analista, custom_colors):
    fig_status = px.pie(
        names=['Subsídio Parcial', 'Fora do Escopo', 'Subsídio Completo'],
        values=[total_parcial_analista, total_fora_analista, total_completo_analista],
        color_discrete_sequence=custom_colors
    )
    fig_status.update_traces(
        hovertemplate='Tarefas %{label} = %{value}<extra></extra>',
    )
    fig_status.update_layout(
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5
        )
    )
    return fig_status

import plotly.express as px

def exibir_grafico_tp_causa(df_analista, analista_selecionado, custom_colors, st):
    """
    Gera e exibe um gráfico de pizza com as quantidades de tarefas cadastradas por "TP CAUSA (TP COMPLEMENTO)"
    para um analista específico.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - custom_colors: Lista de cores personalizadas para o gráfico.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """
    if 'TP CAUSA (TP COMPLEMENTO)' in df_analista.columns:
        # Filtrar apenas as tarefas cadastradas
        df_cadastradas = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']

        # Contar a quantidade de tarefas por "TP CAUSA (TP COMPLEMENTO)"
        tp_causa_counts = df_cadastradas['TP CAUSA (TP COMPLEMENTO)'].dropna().value_counts().reset_index()
        tp_causa_counts.columns = ['TP Causa', 'Quantidade']

        # Criar o gráfico de pizza
        fig_tp_causa = px.pie(
            names=tp_causa_counts['TP Causa'],
            values=tp_causa_counts['Quantidade'],
            color_discrete_sequence=custom_colors
        )

        # Personalizar o hover e layout do gráfico
        fig_tp_causa.update_traces(
            hovertemplate='Causa %{label} = %{value}<extra></extra>',
        )

        fig_tp_causa.update_layout(
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.1,
                xanchor="center",
                x=0.5
            )
        )

        # Exibir o gráfico na dashboard
        st.plotly_chart(fig_tp_causa, use_container_width=True)
    else:
        st.write("A coluna 'TP CAUSA (TP COMPLEMENTO)' não foi encontrada no dataframe.")


def exibir_grafico_filas_realizadas(df_analista, analista_selecionado, custom_colors, st):
    """
    Gera e exibe um gráfico de pizza com as filas realizadas por um analista específico.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - custom_colors: Lista de cores personalizadas para o gráfico.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    ""
    """

    if 'FILA' in df_analista.columns:
        # Contar a quantidade de tarefas por fila
        filas_feitas_analista = df_analista['FILA'].dropna().value_counts().reset_index()
        filas_feitas_analista.columns = ['Tarefa', 'Quantidade']

        # Criar o gráfico de pizza
        fig_filas_feitas_analista = px.pie(
            names=filas_feitas_analista['Tarefa'],
            values=filas_feitas_analista['Quantidade'],
            color_discrete_sequence=custom_colors
        )

        # Personalizar o hover e layout do gráfico
        fig_filas_feitas_analista.update_traces(
            hovertemplate='Tarefas %{label} = %{value}<extra></extra>',
        )

        fig_filas_feitas_analista.update_layout(
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.1,
                xanchor="center",
                x=0.5
            )
        )

        # Exibir o gráfico na dashboard
        st.plotly_chart(fig_filas_feitas_analista)
    else:
        st.write("A coluna 'FILA' não foi encontrada no dataframe.")
        
def format_timedelta_Chart(td):
    """
    Formata um objeto Timedelta em uma string no formato 'X min Y s'.

    Parâmetros:
        - td: Timedelta a ser formatado.
    """
    total_seconds = td.total_seconds()
    minutes = int(total_seconds // 60)
    seconds = int(total_seconds % 60)
    return f"{minutes} min {seconds}s"


import plotly.express as px
import pandas as pd

def exibir_grafico_tmo_por_dia(df_analista, analista_selecionado, calcular_tmo_por_dia, custom_colors, st):
    """
    Gera e exibe um gráfico de barras com o Tempo Médio Operacional (TMO) por dia para um analista específico.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - calcular_tmo_por_dia: Função que calcula o TMO por dia.
        - custom_colors: Lista de cores personalizadas para o gráfico.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """

    # Calcular o TMO por dia
    df_tmo_analista = calcular_tmo_por_dia(df_analista)

    # Verificar se o DataFrame está vazio
    if df_tmo_analista.empty:
        st.warning(f"Não há dados disponíveis para o analista {analista_selecionado}.")
        return

    # Garantir que a coluna 'TMO' está preenchida
    if 'TMO' not in df_tmo_analista.columns or df_tmo_analista['TMO'].isna().all():
        st.warning("Não há valores válidos de TMO disponíveis.")
        return

    # Preencher valores NaN com um timedelta de 0 segundos para evitar erros
    df_tmo_analista['TMO'] = df_tmo_analista['TMO'].fillna(pd.Timedelta(seconds=0))

    # Converter a coluna "TMO" (Timedelta) para minutos e segundos
    df_tmo_analista['TMO_segundos'] = df_tmo_analista['TMO'].dt.total_seconds()
    df_tmo_analista['TMO_minutos'] = df_tmo_analista['TMO_segundos'] / 60

    # Formatar TMO para exibição como "HH:MM:SS"
    df_tmo_analista['TMO_formatado'] = df_tmo_analista['TMO'].apply(format_timedelta_Chart)

    # Verificar se a coluna 'Dia' existe e contém dados válidos
    if 'Dia' not in df_tmo_analista.columns or df_tmo_analista['Dia'].isna().all():
        st.warning("Não há dados de datas disponíveis.")
        return

    # Remover valores NaN na coluna 'Dia'
    df_tmo_analista = df_tmo_analista.dropna(subset=['Dia'])

    # Determinar a data mínima e máxima do dataset
    if df_tmo_analista.empty:
        st.warning("Não há dados suficientes para exibir o gráfico.")
        return

    data_minima = df_tmo_analista['Dia'].min()
    data_maxima = df_tmo_analista['Dia'].max()

    # Criar um slider interativo para seleção de período no formato Dia Mês Ano
    periodo_selecionado_tmo = st.slider(
        "Selecione o período",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=10), data_maxima),
        format="DD MMM YYYY"  # Formato: Dia Mês Ano (ex: 01 Mar 2025)
    )

    # Filtrar os dados com base no período selecionado
    df_tmo_analista = df_tmo_analista[
        (df_tmo_analista['Dia'] >= periodo_selecionado_tmo[0]) & 
        (df_tmo_analista['Dia'] <= periodo_selecionado_tmo[1])
    ]

    # Se após o filtro não houver dados, exibir uma mensagem amigável
    if df_tmo_analista.empty:
        st.warning("Não há registros de TMO para o período selecionado.")
        return

    # Criar o gráfico de barras
    fig_tmo_analista = px.bar(
        df_tmo_analista, 
        x='Dia', 
        y='TMO_minutos', 
        labels={'TMO_minutos': 'TMO (min)', 'Dia': 'Data'},
        text=df_tmo_analista['TMO_formatado'],  # Exibe o tempo formatado nas barras
        color_discrete_sequence=custom_colors
    )
    
    fig_tmo_analista.update_layout(
        xaxis=dict(
            tickvals=df_tmo_analista['Dia'],
            ticktext=[f"{dia.day} {dia.strftime('%b')} {dia.year}" for dia in df_tmo_analista['Dia']],
            title='Data'
        ),
        yaxis=dict(title='TMO (min)'),
        bargap=0.2  # Espaçamento entre as barras
    )

    # Personalizar o gráfico
    fig_tmo_analista.update_traces(
        hovertemplate='Data: %{x}<br>TMO: %{text}',  # Formato do hover
        textfont_color='white'  # Define a cor do texto como branco
    )

    # Exibir o gráfico na dashboard
    st.plotly_chart(fig_tmo_analista, use_container_width=True)

def exibir_grafico_quantidade_por_dia(df_analista, analista_selecionado, custom_colors, st):
    """
    Gera e exibe um gráfico de barras com a quantidade de tarefas realizadas por dia para um analista específico,
    mostrando por padrão apenas os últimos 30 dias e permitindo ajuste de período via slider.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - custom_colors: Lista de cores personalizadas para o gráfico.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """

    # Ordenar os dados por data
    df_analista = df_analista.sort_values(by='DATA DE CONCLUSÃO DA TAREFA')

    # Agrupar os dados por dia e contar a quantidade de tarefas realizadas
    df_quantidade_analista = df_analista.groupby(df_analista['DATA DE CONCLUSÃO DA TAREFA'].dt.date).size().reset_index(name='Quantidade')
    df_quantidade_analista = df_quantidade_analista.rename(columns={'DATA DE CONCLUSÃO DA TAREFA': 'Dia'})

    # Determinar a data mínima e máxima do dataset
    data_minima = df_quantidade_analista['Dia'].min()
    data_maxima = df_quantidade_analista['Dia'].max()

    # Criar um slider interativo para seleção de período no formato Dia Mês Ano
    periodo_selecionado = st.slider(
        "Selecione o período",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=11), data_maxima),
        format="DD MMM YYYY"  # Formato: Dia Mês Ano (ex: 01 Mar 2025)
    )

    # Filtrar os dados com base no período selecionado
    df_quantidade_analista = df_quantidade_analista[
        (df_quantidade_analista['Dia'] >= periodo_selecionado[0]) &
        (df_quantidade_analista['Dia'] <= periodo_selecionado[1])
    ]

    # Criar o gráfico de barras
    fig_quantidade_analista = px.bar(
        df_quantidade_analista, 
        x='Dia', 
        y='Quantidade', 
        labels={'Quantidade': 'Quantidade de Tarefas', 'Dia': 'Data'},
        text='Quantidade',  # Exibe a quantidade nas barras
        color_discrete_sequence=custom_colors
    )

    # Ajuste para melhorar a legibilidade
    fig_quantidade_analista.update_layout(
        xaxis=dict(
            tickvals=df_quantidade_analista['Dia'],
            ticktext=[f"{dia.day} {dia.strftime('%b')} {dia.year}" for dia in df_quantidade_analista['Dia']],  # Formato Dia Mês Ano
            title='Data'
        ),
        yaxis=dict(title='Quantidade de Tarefas'),
        bargap=0.2  # Espaçamento entre as barras
    )

    # Personalizar o gráfico
    fig_quantidade_analista.update_traces(
        hovertemplate='Data: %{x}<br>Quantidade: %{y}',  # Formato do hover
        textfont_color='white'  # Define a cor do texto como branco
    )

    # Exibir o gráfico na dashboard
    st.plotly_chart(fig_quantidade_analista, use_container_width=True)

def plot_grafico_desvios(df):
    if 'DESVIOS CADASTRO' not in df.columns:
        return None

    desvios_series = df['DESVIOS CADASTRO'].dropna().str.split(',')
    desvios_explodidos = desvios_series.explode().str.strip()
    desvios_filtrados = desvios_explodidos[desvios_explodidos.str.lower() != "sem desvio"]

    contagem = desvios_filtrados.value_counts().reset_index()
    contagem.columns = ['Desvio', 'Frequência']

    fig = px.bar(
        contagem,
        x='Desvio',
        y='Frequência',
        color='Desvio',
        title='Desvios Encontrados - Gráfico de Barras Coloridas'
    )
    fig.update_layout(xaxis_tickangle=-45)
    return fig
