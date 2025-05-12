import pandas as pd
import os
import plotly.express as px
import math
import streamlit as st
from io import BytesIO
from datetime import timedelta

def load_data(usuario):
    parquet_file = f'dados_acumulados_{usuario}.parquet'  # Caminho do arquivo Parquetaa
    
    try:
        if os.path.exists(parquet_file):
            df_total = pd.read_parquet(parquet_file)
            
            # Verifica se a coluna 'Justificativa' existe, caso contrário, adiciona ela
            if 'Justificativa' not in df_total.columns:
                df_total['Justificativa'] = ""  # Adiciona a coluna com valores vazios
        else:
            raise FileNotFoundError
    
    except (FileNotFoundError, ValueError, OSError):
        # Cria um DataFrame vazio com a coluna 'Justificativa' e salva um novo arquivo
        df_total = pd.DataFrame(columns=[
            'NÚMERO DO PROTOCOLO', 
            'USUÁRIO QUE CONCLUIU A TAREFA', 
            'SITUAÇÃO DA TAREFA', 
            'TEMPO MÉDIO OPERACIONAL', 
            'DATA DE CONCLUSÃO DA TAREFA', 
            'FINALIZAÇÃO',
            'Justificativa'  # Inclui a coluna de justificativa ao criar um novo DataFrame
        ])
        df_total.to_parquet(parquet_file, index=False)
    
    return df_total

def save_data(df, usuario):
    import os
    parquet_file = f'dados_acumulados_{usuario}.parquet'
    log_file = f'log_ajustes_tmo_{usuario}.csv'
    ajustes = []

    # Remove colunas desnecessárias
    df = df.loc[:, ~(df.columns.str.upper().str.strip().isin(['ID NIP', 'M.O.', 'Nº LB (JV - CÍVEL)', 'Nº LB (AMIL - CÍVEL)', 'Nº LB (JV TRABALHISTA)']))]

    # Garante que a coluna 'Justificativa' exista
    if 'Justificativa' not in df.columns:
        df['Justificativa'] = ""

    # Remove registros automáticos
    if 'USUÁRIO QUE CONCLUIU A TAREFA' in df.columns:
        df = df[
            (df['USUÁRIO QUE CONCLUIU A TAREFA'].notna()) &
            (df['USUÁRIO QUE CONCLUIU A TAREFA'].str.lower() != 'robohub_amil')
        ]

    # 🔄 Padronizações de TMO
    if 'TEMPO MÉDIO OPERACIONAL' in df.columns and 'FINALIZAÇÃO' in df.columns:
        df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')

        for i, row in df.iterrows():
            tmo = row['TEMPO MÉDIO OPERACIONAL']
            finalizacao = row['FINALIZAÇÃO']
            protocolo = row.get('NÚMERO DO PROTOCOLO', 'N/A')
            novo_tmo = tmo

            if finalizacao == 'CADASTRADO' and pd.notnull(tmo) and tmo < pd.Timedelta(minutes=19):
                novo_tmo = pd.Timedelta(minutes=20)
            elif finalizacao == 'ATUALIZADO':
                if pd.notnull(tmo) and tmo < pd.Timedelta(minutes=3):
                    novo_tmo = pd.Timedelta(minutes=3)
                elif pd.notnull(tmo) and tmo > pd.Timedelta(minutes=15):
                    novo_tmo = pd.Timedelta(minutes=15)

            if pd.notnull(tmo) and tmo > pd.Timedelta(hours=2):
                novo_tmo = pd.Timedelta(hours=2)

            if pd.notnull(tmo) and novo_tmo != tmo:
                ajustes.append({
                    'NÚMERO DO PROTOCOLO': protocolo,
                    'FINALIZAÇÃO': finalizacao,
                    'TMO ORIGINAL': tmo,
                    'TMO AJUSTADO': novo_tmo
                })
                df.at[i, 'TEMPO MÉDIO OPERACIONAL'] = novo_tmo

    # Salva o DataFrame atualizado
    df.to_parquet(parquet_file, index=False)

    # Salva log: se houver ajustes → CSV detalhado | senão → mensagem simples
    if ajustes:
        df_ajustes = pd.DataFrame(ajustes)
        df_ajustes.to_csv(log_file, index=False)
    else:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write("Nenhum ajuste de TMO foi necessário.\n")

    return df

def calcular_tmo_por_dia(df):
    df['Dia'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA']).dt.date
    df_finalizados = df[df['SITUAÇÃO DA TAREFA'].isin(['Finalizada', 'Cancelada'])].copy()
    
    # Agrupando por dia
    df_tmo = df_finalizados.groupby('Dia').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),  # Soma total do tempo
        Total_Finalizados_Cancelados=('SITUAÇÃO DA TAREFA', 'count')  # Total de tarefas finalizadas ou canceladas
    ).reset_index()

    # Calcula o TMO (Tempo Médio Operacional)
    df_tmo['TMO'] = df_tmo['Tempo_Total'] / df_tmo['Total_Finalizados_Cancelados']
    
    # Formata o tempo médio no formato HH:MM:SS
    df_tmo['TMO'] = df_tmo['TMO'].apply(format_timedelta)
    return df_tmo[['Dia', 'TMO']]

def calcular_tmo_por_dia_geral(df):
    # Certifica-se de que a coluna de data está no formato correto
    df['Dia'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA']).dt.date

    # Filtra tarefas finalizadas ou canceladas, pois estas são relevantes para o cálculo do TMO
    df_finalizados = df[df['SITUAÇÃO DA TAREFA'].isin(['Finalizado', 'Cancelada'])].copy()
    
    # Agrupamento por dia para calcular o tempo médio diário
    df_tmo = df_finalizados.groupby('Dia').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),  # Soma total do tempo por dia
        Total_Finalizados_Cancelados=('SITUAÇÃO DA TAREFA', 'count')  # Total de tarefas finalizadas/canceladas por dia
    ).reset_index()

    # Calcula o TMO (Tempo Médio Operacional) diário
    df_tmo['TMO'] = df_tmo['Tempo_Total'] / df_tmo['Total_Finalizados_Cancelados']
    
    # Remove valores nulos e formata o tempo médio para o gráfico
    df_tmo['TMO'] = df_tmo['TMO'].fillna(pd.Timedelta(seconds=0))  # Preenche com zero se houver NaN
    df_tmo['TMO_Formatado'] = df_tmo['TMO'].apply(format_timedelta)  # Formata para exibição
    
    return df_tmo[['Dia', 'TMO', 'TMO_Formatado']]

def calcular_produtividade_diaria(df):
    # Garante que a coluna 'Próximo' esteja em formato de data
    df['Dia'] = df['DATA DE CONCLUSÃO DA TAREFA'].dt.date

    # Agrupa e soma os status para calcular a produtividade
    df_produtividade = df.groupby('Dia').agg(
        Finalizado=('FINALIZAÇÃO', 'count'),
    ).reset_index()

    # Calcula a produtividade total
    df_produtividade['Produtividade'] = + df_produtividade['Finalizado'] 
    return df_produtividade

def calcular_produtividade_diaria_cadastro(df):
    # Garante que a coluna 'Próximo' esteja em formato de data
    df['Dia'] = df['DATA DE CONCLUSÃO DA TAREFA'].dt.date

    # Agrupa e soma os status para calcular a produtividade
    df_produtividade_cadastro = df.groupby('Dia').agg(
        Finalizado=('FINALIZAÇÃO', lambda x: x[x == 'CADASTRADO'].count()),
        Atualizado=('FINALIZAÇÃO', lambda x: x[x == 'ATUALIZADO'].count())
    ).reset_index()

    # Calcula a produtividade total
    df_produtividade_cadastro['Produtividade'] = + df_produtividade_cadastro['Finalizado'] + df_produtividade_cadastro['Atualizado']
    return df_produtividade_cadastro

def convert_to_timedelta_for_calculations(df):
    df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')
    return df

def convert_to_datetime_for_calculations(df):
    df['DATA DE CONCLUSÃO DA TAREFA'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    return df
        
def format_timedelta(td):
    if pd.isnull(td):
        return "0 min"
    total_seconds = int(td.total_seconds())
    minutes, seconds = divmod(total_seconds, 60)
    return f"{minutes} min {seconds}s"

def format_timedelta_grafico_tmo(td):
    if pd.isnull(td):
        return "00:00:00"
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

# Função para calcular o TMO por analista
def calcular_tmo_por_dia(df):
    df['Dia'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA']).dt.date
    df_finalizados = df[df['SITUAÇÃO DA TAREFA'].isin(['Finalizada', 'Cancelada'])].copy()
    
    # Agrupando por dia
    df_tmo = df_finalizados.groupby('Dia').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),  # Soma total do tempo
        Total_Finalizados_Cancelados=('SITUAÇÃO DA TAREFA', 'count')  # Total de tarefas finalizadas ou canceladas
    ).reset_index()

    # Calcula o TMO (Tempo Médio Operacional)
    df_tmo['TMO'] = df_tmo['Tempo_Total'] / df_tmo['Total_Finalizados_Cancelados']
    
    # Formata o tempo médio no formato HH:MM:SS
    df_tmo['TMO'] = df_tmo['TMO'].apply(format_timedelta)
    return df_tmo[['Dia', 'TMO']]

def calcular_tmo_por_dia_cadastro(df):
    df['Dia'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA']).dt.date
    df_finalizados_cadastro = df[df['FINALIZAÇÃO'] == 'CADASTRADO'].copy()
    
    # Agrupando por dia
    df_tmo_cadastro = df_finalizados_cadastro.groupby('Dia').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),  # Soma total do tempo
        Total_Finalizados_Cancelados=('FINALIZAÇÃO', 'count')  # Total de tarefas finalizadas ou canceladas
    ).reset_index()

    # Calcula o TMO (Tempo Médio Operacional)
    df_tmo_cadastro['TMO'] = df_tmo_cadastro['Tempo_Total'] / df_tmo_cadastro['Total_Finalizados_Cancelados']
    
    # Formata o tempo médio no formato HH:MM:SS
    df_tmo_cadastro['TMO'] = df_tmo_cadastro['TMO'].apply(format_timedelta)
    return df_tmo_cadastro[['Dia', 'TMO']]

# Função para calcular o TMO por analista
def calcular_tmo(df):
    # Verifica se a coluna 'SITUAÇÃO DA TAREFA' existe no DataFrame
    if 'SITUAÇÃO DA TAREFA' not in df.columns:
        raise KeyError("A coluna 'SITUAÇÃO DA TAREFA' não foi encontrada no DataFrame.")

    # Filtra as tarefas finalizadas ou canceladas
    df_finalizados = df[df['SITUAÇÃO DA TAREFA'].isin(['Finalizada', 'Cancelada'])].copy()

    # Verifica se a coluna 'TEMPO MÉDIO OPERACIONAL' existe e converte para minutos
    if 'TEMPO MÉDIO OPERACIONAL' not in df_finalizados.columns:
        raise KeyError("A coluna 'TEMPO MÉDIO OPERACIONAL' não foi encontrada no DataFrame.")
    df_finalizados['TEMPO_MÉDIO_MINUTOS'] = df_finalizados['TEMPO MÉDIO OPERACIONAL'].dt.total_seconds() / 60

    # Verifica se a coluna 'FILA' existe antes de aplicar o filtro
    if 'FILA' in df_finalizados.columns:
        # Remove protocolos da fila "DÚVIDA" com mais de 1 hora de tempo médio
        df_finalizados = df_finalizados[~((df_finalizados['FILA'] == 'DÚVIDA') & (df_finalizados['TEMPO_MÉDIO_MINUTOS'] > 60))]

    # Agrupando por analista
    df_tmo_analista = df_finalizados.groupby('USUÁRIO QUE CONCLUIU A TAREFA').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', lambda x: x[df_finalizados['FINALIZAÇÃO'] == 'CADASTRADO'].sum()),  # Soma total do tempo das tarefas com finalização CADASTRADO
        Total_Tarefas=('FINALIZAÇÃO', lambda x: x[x == 'CADASTRADO'].count())  # Total de tarefas finalizadas ou canceladas por analista
    ).reset_index()

    # Calcula o TMO (Tempo Médio Operacional) como média
    df_tmo_analista['TMO'] = df_tmo_analista['Tempo_Total'] / df_tmo_analista['Total_Tarefas']

    # Formata o tempo médio no formato de minutos e segundos
    df_tmo_analista['TMO_Formatado'] = df_tmo_analista['TMO'].apply(format_timedelta_grafico_tmo)

    return df_tmo_analista[['USUÁRIO QUE CONCLUIU A TAREFA', 'TMO_Formatado', 'TMO']]

# Função para calcular o ranking dinâmico
def calcular_ranking(df_total, selected_users):
    # Filtra o DataFrame com os usuários selecionados
    df_filtered = df_total[df_total['USUÁRIO QUE CONCLUIU A TAREFA'].isin(selected_users)]

    df_ranking = df_filtered.groupby('USUÁRIO QUE CONCLUIU A TAREFA').agg(

        Finalizado=('FINALIZAÇÃO', lambda x: x[x == 'CADASTRADO'].count()),
        Distribuido=('FINALIZAÇÃO', lambda x: x[x == 'REALIZADO'].count()),
        Atualizado=('FINALIZAÇÃO', lambda x: x[x == 'ATUALIZADO'].count())
    ).reset_index()
    df_ranking['Total'] =df_ranking['Finalizado'] + df_ranking['Distribuido'] + df_ranking['Atualizado']
    df_ranking = df_ranking.sort_values(by  ='Total', ascending=False).reset_index(drop=True)
    df_ranking.index += 1
    df_ranking.index.name = 'Posição'

    # Define o tamanho dos quartis
    num_analistas = len(df_ranking)
    quartil_size = 4 if num_analistas > 12 else math.ceil(num_analistas / 4)

    # Função de estilo para os quartis dinâmicos
    def apply_dynamic_quartile_styles(row):
        if row.name <= quartil_size:
            color = 'rgba(135, 206, 250, 0.4)'  # Azul vibrante translúcido (primeiro quartil)
        elif quartil_size < row.name <= 2 * quartil_size:
            color = 'rgba(144, 238, 144, 0.4)'  # Verde vibrante translúcido (segundo quartil)
        elif 2 * quartil_size < row.name <= 3 * quartil_size:
            color = 'rgba(255, 255, 102, 0.4)'  # Amarelo vibrante translúcido (terceiro quartil)
        else:
            color = 'rgba(255, 99, 132, 0.4)'  # Vermelho vibrante translúcido (quarto quartil)
        return ['background-color: {}'.format(color) for _ in row]

    # Aplicar os estilos e retornar o DataFrame
    styled_df_ranking = df_ranking.style.apply(apply_dynamic_quartile_styles, axis=1).format(
        {'Andamento': '{:.0f}', 'Finalizado': '{:.0f}', 'Reclassificado': '{:.0f}', 'Total': '{:.0f}'}
    )

    return styled_df_ranking

def obter_melhor_analista_por_fila(df):
    df = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO'])].copy()
    df = df[~df['USUÁRIO QUE CONCLUIU A TAREFA'].str.contains('_ter', na=False)]
    df = df[df['USUÁRIO QUE CONCLUIU A TAREFA'] != 'viniciusgimenes_amil']
    df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')
    agrupado = df.groupby(['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA'])
    resultado = agrupado.agg(
        TMO=('TEMPO MÉDIO OPERACIONAL', 'mean'),
        Quantidade=('TEMPO MÉDIO OPERACIONAL', 'count')
    ).reset_index()

    melhores = resultado.loc[resultado.groupby('FILA')['TMO'].idxmin()].reset_index(drop=True)
    melhores['TMO'] = melhores['TMO'].apply(lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}")
    return melhores

def obter_maior_quantidade_por_fila(df):
    colunas_necessarias = {'FINALIZAÇÃO', 'USUÁRIO QUE CONCLUIU A TAREFA', 'FILA'}
    if not colunas_necessarias.issubset(df.columns):
        return pd.DataFrame(columns=['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA', 'Quantidade'])

    df = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO'])].copy()
    df = df[~df['USUÁRIO QUE CONCLUIU A TAREFA'].str.contains("_ter|viniciusgimenes_amil", na=False)]

    if df.empty:
        return pd.DataFrame(columns=['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA', 'Quantidade'])

    resultado = df.groupby(['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA']).size().reset_index(name='Quantidade')
    maiores = resultado.loc[resultado.groupby('FILA')['Quantidade'].idxmax()].reset_index(drop=True)
    return maiores

def exibir_maior_quantidade_por_fila(df):
    st.subheader("Melhor Analista por Fila (Quantidade)")
    maiores = obter_maior_quantidade_por_fila(df)
    if maiores.empty:
        st.info("Ainda não há dados para exibir o analista com maior quantidade por fila.")
    else:
        st.dataframe(maiores, hide_index=True, use_container_width=True)

def obter_melhor_analista_por_fila(df):
    colunas_necessarias = {'FINALIZAÇÃO', 'USUÁRIO QUE CONCLUIU A TAREFA', 'FILA', 'TEMPO MÉDIO OPERACIONAL'}
    if not colunas_necessarias.issubset(df.columns):
        return pd.DataFrame(columns=['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA', 'TMO', 'Quantidade'])

    df = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO'])].copy()
    df = df[~df['USUÁRIO QUE CONCLUIU A TAREFA'].str.contains('_ter', na=False)]
    df = df[df['USUÁRIO QUE CONCLUIU A TAREFA'] != 'viniciusgimenes_amil']
    df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')

    if df.empty:
        return pd.DataFrame(columns=['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA', 'TMO', 'Quantidade'])

    agrupado = df.groupby(['FILA', 'USUÁRIO QUE CONCLUIU A TAREFA'])
    resultado = agrupado.agg(
        TMO=('TEMPO MÉDIO OPERACIONAL', 'mean'),
        Quantidade=('TEMPO MÉDIO OPERACIONAL', 'count')
    ).reset_index()

    melhores = resultado.loc[resultado.groupby('FILA')['TMO'].idxmin()].reset_index(drop=True)
    melhores['TMO'] = melhores['TMO'].apply(lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}")
    return melhores

def exibir_melhor_analista_por_fila(df):
    st.subheader("Melhor Analista por Fila (TMO)")
    melhores = obter_melhor_analista_por_fila(df)
    if melhores.empty:
        st.info("Ainda não há dados para exibir o melhor analista por fila.")
    else:
        st.dataframe(melhores, hide_index=True, use_container_width=True)

    
def calcular_cadastro_atualizacao_por_modulo(df):
    # Verifica se as colunas necessárias existem
    if 'MÓDULO LB' not in df.columns or 'FINALIZAÇÃO' not in df.columns:
        return pd.DataFrame(columns=['MÓDULO LB', 'CADASTRADO', 'ATUALIZADO'])

    # Filtra apenas CADASTRADO e ATUALIZADO
    df = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO'])].copy()

    # Agrupa e conta
    df_resultado = df.groupby(['MÓDULO LB', 'FINALIZAÇÃO']).size().unstack(fill_value=0).reset_index()
    df_resultado.columns.name = None  # remove nome do índice de coluna
    return df_resultado

def exibir_cadastro_atualizacao_por_modulo(df):
    st.subheader("Quantidade de Cadastros e Atualizações por Módulo")

    df_modulo = calcular_cadastro_atualizacao_por_modulo(df)

    if df_modulo.empty:
        st.info("Ainda não há dados para exibir cadastros e atualizações por módulo.")
        return

    st.dataframe(df_modulo, use_container_width=True, hide_index=True)

#MÉTRICAS INDIVIDUAIS
import pandas as pd
import streamlit as st

def calcular_metrica_analista(df_analista):
    # Verifica se as colunas necessárias estão presentes no DataFrame
    colunas_necessarias = ['FILA', 'FINALIZAÇÃO', 'TEMPO MÉDIO OPERACIONAL', 'DATA DE CONCLUSÃO DA TAREFA']
    for coluna in colunas_necessarias:
        if coluna not in df_analista.columns:
            st.warning(f"A coluna '{coluna}' não está disponível nos dados. Verifique o arquivo carregado.")
            return None, None, None, None, None, None, None  # Atualizado para retornar sete valores

    # Excluir os registros com "FILA" como "Desconhecida"
    df_analista_filtrado = df_analista[df_analista['FILA'] != "Desconhecida"]

    # Filtrar os registros com status "CADASTRADO", "ATUALIZADO" e "REALIZADO"
    df_filtrados = df_analista_filtrado[df_analista_filtrado['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO', 'REALIZADO'])]

    # Converter "TEMPO MÉDIO OPERACIONAL" para minutos
    df_filtrados['TEMPO_MÉDIO_MINUTOS'] = df_filtrados['TEMPO MÉDIO OPERACIONAL'].dt.total_seconds() / 60

    # Excluir registros da fila "DÚVIDA" com tempo médio superior a 1 hora
    df_filtrados = df_filtrados[~((df_filtrados['FILA'] == 'DÚVIDA') & (df_filtrados['TEMPO_MÉDIO_MINUTOS'] > 60))]

    # Calcula totais conforme os filtros de status
    total_finalizados = len(df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'CADASTRADO'])
    total_realizados = len(df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'REALIZADO'])
    total_atualizado = len(df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'ATUALIZADO'])

    # Calcula o tempo total para cada tipo de tarefa
    tempo_total_cadastrado = df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'CADASTRADO']['TEMPO MÉDIO OPERACIONAL'].sum()
    tempo_total_atualizado = df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'ATUALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum()
    tempo_total_realizado = df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'REALIZADO']['TEMPO MÉDIO OPERACIONAL'].sum()

    # Calcula o tempo médio para cada tipo de tarefa
    tmo_cadastrado = tempo_total_cadastrado / total_finalizados if total_finalizados > 0 else pd.Timedelta(0)
    tmo_atualizado = tempo_total_atualizado / total_atualizado if total_atualizado > 0 else pd.Timedelta(0)

    # Calcula o tempo médio geral considerando todas as tarefas
    tempo_total_analista = tempo_total_cadastrado + tempo_total_atualizado + tempo_total_realizado
    total_tarefas = total_finalizados + total_atualizado + total_realizados
    tempo_medio_analista = tempo_total_analista / total_tarefas if total_tarefas > 0 else pd.Timedelta(0)

    # Calcular a média de cadastros por dias trabalhados
    dias_trabalhados = df_filtrados[df_filtrados['FINALIZAÇÃO'] == 'CADASTRADO']['DATA DE CONCLUSÃO DA TAREFA'].dt.date.nunique()
    media_cadastros_por_dia = int(total_finalizados / dias_trabalhados) if dias_trabalhados > 0 else 0

    return total_finalizados, total_atualizado, tempo_medio_analista, tmo_cadastrado, tmo_atualizado, total_realizados, media_cadastros_por_dia, dias_trabalhados

def calcular_tempo_ocioso_por_analista(df):
    try:
        # Converte as colunas de data para datetime, tratando erros
        df['DATA DE INÍCIO DA TAREFA'] = pd.to_datetime(
            df['DATA DE INÍCIO DA TAREFA'], format='%d/%m/%Y %H:%M:%S', errors='coerce'
        )
        df['DATA DE CONCLUSÃO DA TAREFA'] = pd.to_datetime(
            df['DATA DE CONCLUSÃO DA TAREFA'], format='%d/%m/%Y %H:%M:%S', errors='coerce'
        )

        # Filtrar valores nulos e resetar index
        df = df.dropna(subset=['DATA DE INÍCIO DA TAREFA', 'DATA DE CONCLUSÃO DA TAREFA']).reset_index(drop=True)

        # Ordena os dados por usuário e data de início da tarefa (sem considerar fila)
        df = df.sort_values(by=['USUÁRIO QUE CONCLUIU A TAREFA', 'DATA DE INÍCIO DA TAREFA']).reset_index(drop=True)

        # Calcula o próximo horário de início da tarefa por usuário (sem considerar fila)
        df['PRÓXIMA_TAREFA'] = df.groupby(['USUÁRIO QUE CONCLUIU A TAREFA'])['DATA DE INÍCIO DA TAREFA'].shift(-1)

        # Calcula o tempo ocioso entre a conclusão de uma tarefa e o início da próxima
        df['TEMPO OCIOSO'] = df['PRÓXIMA_TAREFA'] - df['DATA DE CONCLUSÃO DA TAREFA']

        # Remove valores negativos ou muito grandes (exemplo: trocas de turno)
        df['TEMPO OCIOSO'] = df['TEMPO OCIOSO'].apply(lambda x: x if pd.notnull(x) and pd.Timedelta(0) < x <= pd.Timedelta(hours=1) else pd.Timedelta(0))

        # Agrupa os tempos ociosos por usuário e dia de conclusão, somando todas as filas
        df_soma_ocioso = df.groupby(['USUÁRIO QUE CONCLUIU A TAREFA', df['DATA DE CONCLUSÃO DA TAREFA'].dt.date])['TEMPO OCIOSO'].sum().reset_index()

        # Renomeia colunas para melhor entendimento
        df_soma_ocioso = df_soma_ocioso.rename(columns={
            'DATA DE CONCLUSÃO DA TAREFA': 'Data',
            'TEMPO OCIOSO': 'Tempo Ocioso'
        })

        # Formata o tempo ocioso total como string
        df_soma_ocioso['Tempo Ocioso'] = df_soma_ocioso['Tempo Ocioso'].astype(str).str.split("days").str[-1].str.strip()

        return df_soma_ocioso[['USUÁRIO QUE CONCLUIU A TAREFA', 'Data', 'Tempo Ocioso']]

    except Exception as e:
        return pd.DataFrame({'Erro': [f'Erro: {str(e)}']})

def format_timedelta_hms(td):
    """ Formata um timedelta para HH:MM:SS """
    if pd.isnull(td):
        return "00:00:00"
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def exibir_grafico_tempo_ocioso_por_dia(df_analista, analista_selecionado, custom_colors, st):
    """
    Gera e exibe um gráfico de barras com o Tempo Ocioso diário para um analista específico.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - custom_colors: Lista de cores personalizadas para o gráfico.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """

    # Calcular o tempo ocioso diário por analista
    df_ocioso = calcular_tempo_ocioso_por_analista(df_analista)

    # Converter a coluna 'Tempo Ocioso' para timedelta para cálculos
    df_ocioso['Tempo Ocioso'] = pd.to_timedelta(df_ocioso['Tempo Ocioso'], errors='coerce')

    # Filtrar apenas o analista selecionado
    df_ocioso = df_ocioso[df_ocioso['USUÁRIO QUE CONCLUIU A TAREFA'] == analista_selecionado]

    # Determinar o período disponível no dataset
    data_minima = df_ocioso['Data'].min()
    data_maxima = df_ocioso['Data'].max()

    # Criar um slider interativo para seleção de período
    periodo_selecionado = st.slider(
        "Selecione o período",
        min_value=data_minima,
        max_value=data_maxima,
        value=(data_maxima - pd.Timedelta(days=30), data_maxima),  # Últimos 30 dias por padrão
        format="DD MMM YYYY"  # Formato: Dia Mês Ano (Exemplo: 01 Jan 2025)
    )

    # Filtrar os dados com base no período selecionado
    df_ocioso = df_ocioso[
        (df_ocioso['Data'] >= periodo_selecionado[0]) &
        (df_ocioso['Data'] <= periodo_selecionado[1])
    ]

    # Formatar a coluna 'Tempo Ocioso' para exibição no gráfico como HH:MM:SS
    df_ocioso['Tempo Ocioso Formatado'] = df_ocioso['Tempo Ocioso'].apply(format_timedelta_hms)

    # Converter tempo ocioso para total de segundos (para exibição correta no gráfico)
    df_ocioso['Tempo Ocioso Segundos'] = df_ocioso['Tempo Ocioso'].dt.total_seconds()

    # Criar o gráfico de barras
    fig_ocioso = px.bar(
        df_ocioso, 
        x='Data', 
        y='Tempo Ocioso Segundos', 
        labels={'Tempo Ocioso Segundos': 'Tempo Ocioso (HH:MM:SS)', 'Data': 'Data'},
        text=df_ocioso['Tempo Ocioso Formatado'],  # Exibir tempo formatado nas barras
        color_discrete_sequence=custom_colors
    )

    # Ajuste do layout
    fig_ocioso.update_layout(
        xaxis=dict(
            tickvals=df_ocioso['Data'],
            ticktext=[f"{dia.day} {dia.strftime('%b')} {dia.year}" for dia in df_ocioso['Data']],
            title='Data'
        ),
        yaxis=dict(
            title='Tempo Ocioso (HH:MM:SS)',
            tickvals=[i * 3600 for i in range(0, int(df_ocioso['Tempo Ocioso Segundos'].max() // 3600) + 1)],
            ticktext=[format_timedelta_hms(pd.Timedelta(seconds=i * 3600)) for i in range(0, int(df_ocioso['Tempo Ocioso Segundos'].max() // 3600) + 1)]
        ),
        bargap=0.2  # Espaçamento entre as barras
    )

    # Personalizar o gráfico
    fig_ocioso.update_traces(
        hovertemplate='Data = %{x}<br>Tempo Ocioso = %{text}',  # Formato do hover
        textposition="outside",  # Exibir rótulos acima das barras
        textfont_color='white'  # Define a cor do texto como branco
    )

    # Exibir o gráfico na dashboard
    st.plotly_chart(fig_ocioso, use_container_width=True)

def calcular_tmo_equipe_cadastro(df_total):
    return df_total[df_total['FINALIZAÇÃO'].isin(['CADASTRADO'])]['TEMPO MÉDIO OPERACIONAL'].mean()

def calcular_tmo_equipe_atualizado(df_total):
    return df_total[df_total['FINALIZAÇÃO'].isin(['ATUALIZADO'])]['TEMPO MÉDIO OPERACIONAL'].mean()

def calcular_filas_analista(df_analista):
    if 'Carteira' in df_analista.columns:
        # Filtra apenas os status relevantes para o cálculo (considerando FINALIZADO e RECLASSIFICADO)
        filas_finalizadas_analista = df_analista[
            df_analista['Status'].isin(['FINALIZADO', 'RECLASSIFICADO', 'ANDAMENTO_PRE'])
        ]
        
        # Agrupa por 'Carteira' e calcula a quantidade de FINALIZADO, RECLASSIFICADO e ANDAMENTO_PRE para cada fila
        carteiras_analista = filas_finalizadas_analista.groupby('Carteira').agg(
            Finalizados=('Status', lambda x: (x == 'FINALIZADO').sum()),
            Reclassificados=('Status', lambda x: (x == 'RECLASSIFICADO').sum()),
            Andamento=('Status', lambda x: (x == 'ANDAMENTO_PRE').sum()),
            TMO_médio=('Tempo de Análise', lambda x: x[x.index.isin(df_analista[(df_analista['Status'].isin(['FINALIZADO', 'RECLASSIFICADO']))].index)].mean())
        ).reset_index()

        # Converte o TMO médio para minutos e segundos
        carteiras_analista['TMO_médio'] = carteiras_analista['TMO_médio'].apply(format_timedelta)

        # Renomeia as colunas para exibição
        carteiras_analista = carteiras_analista.rename(
            columns={'Carteira': 'Fila', 'Finalizados': 'Finalizados', 'Reclassificados': 'Reclassificados', 'Andamento': 'Andamento', 'TMO_médio': 'TMO Médio por Fila'}
        )
        
        return carteiras_analista  # Retorna o DataFrame
    
    else:
        # Caso a coluna 'Carteira' não exista
        return pd.DataFrame({'Fila': [], 'Finalizados': [], 'Reclassificados': [], 'Andamento': [], 'TMO Médio por Fila': []})
    


def calcular_tmo_por_dia(df_analista):
    # Filtrar apenas as tarefas com finalização "CADASTRADO"
    df_analista = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']
    df_analista['Dia'] = df_analista['DATA DE CONCLUSÃO DA TAREFA'].dt.date
    tmo_por_dia = df_analista.groupby('Dia').agg(TMO=('TEMPO MÉDIO OPERACIONAL', 'mean')).reset_index()
    return tmo_por_dia

def calcular_carteiras_analista(df_analista):
    if 'Carteira' in df_analista.columns:
        filas_finalizadas = df_analista[(df_analista['Status'] == 'FINALIZADO') |
                                        (df_analista['Status'] == 'RECLASSIFICADO') |
                                        (df_analista['Status'] == 'ANDAMENTO_PRE')]

        carteiras_analista = filas_finalizadas.groupby('Carteira').agg(
            Quantidade=('Carteira', 'size'),
            TMO_médio=('Tempo de Análise', 'mean')
        ).reset_index()

        # Renomeando a coluna 'Carteira' para 'Fila' para manter consistência
        carteiras_analista = carteiras_analista.rename(columns={'Carteira': 'Fila'})

        return carteiras_analista
    else:
        return pd.DataFrame({'Fila': [], 'Quantidade': [], 'TMO Médio por Fila': []})
    
def get_points_of_attention(df):
    # Verifica se a coluna 'Carteira' existe no DataFrame
    if 'Carteira' not in df.columns:
        return "A coluna 'Carteira' não foi encontrada no DataFrame."
    
    # Filtra os dados para 'JV ITAU BMG' e outras carteiras
    dfJV = df[df['Carteira'] == 'JV ITAU BMG'].copy()
    dfOutras = df[df['Carteira'] != 'JV ITAU BMG'].copy()
    
    # Filtra os pontos de atenção com base no tempo de análise
    pontos_de_atencao_JV = dfJV[dfJV['Tempo de Análise'] > pd.Timedelta(minutes=2)]
    pontos_de_atencao_outros = dfOutras[dfOutras['Tempo de Análise'] > pd.Timedelta(minutes=5)]
    
    # Combina os dados filtrados
    pontos_de_atencao = pd.concat([pontos_de_atencao_JV, pontos_de_atencao_outros])

    # Verifica se o DataFrame está vazio
    if pontos_de_atencao.empty:
        return "Não existem dados a serem exibidos."

    # Cria o dataframe com as colunas 'PROTOCOLO', 'CARTEIRA' e 'TEMPO'
    pontos_de_atencao = pontos_de_atencao[['Protocolo', 'Carteira', 'Tempo de Análise']].copy()

    # Renomeia a coluna 'Tempo de Análise' para 'TEMPO'
    pontos_de_atencao = pontos_de_atencao.rename(columns={'Tempo de Análise': 'TEMPO'})

    # Converte a coluna 'TEMPO' para formato de minutos
    pontos_de_atencao['TEMPO'] = pontos_de_atencao['TEMPO'].apply(lambda x: f"{int(x.total_seconds() // 60)}:{int(x.total_seconds() % 60):02d}")

    # Remove qualquer protocolo com valores vazios ou NaN
    pontos_de_atencao = pontos_de_atencao.dropna(subset=['Protocolo'])

    # Remove as vírgulas e a parte ".0" do protocolo
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].astype(str).str.replace(',', '', regex=False)
    
    # Garantir que o número do protocolo não tenha ".0"
    pontos_de_atencao['Protocolo'] = pontos_de_atencao['Protocolo'].str.replace(r'\.0$', '', regex=True)

    return pontos_de_atencao

import pandas as pd

import pandas as pd

def calcular_tmo_por_carteira(df):
    """
    Calcula o Tempo Médio Operacional (TMO) por carteira, separando o TMO de Cadastro e o TMO de Atualização.
    Para a fila "Distribuição", o cálculo do TMO considera todas as tarefas finalizadas como "REALIZADO".

    Parâmetros:
        df (pd.DataFrame): DataFrame contendo as colunas 'FILA', 'TEMPO MÉDIO OPERACIONAL', 
                           'FINALIZAÇÃO' e 'NÚMERO DO PROTOCOLO'.

    Retorno:
        pd.DataFrame: DataFrame contendo as colunas 'FILA', 'Quantidade', 'Cadastrado', 'Atualizado',
                      'Fora do Escopo', 'TMO Cadastro' e 'TMO Atualização'.
    """
    # Verifica se as colunas necessárias estão no DataFrame
    required_columns = {'FILA', 'TEMPO MÉDIO OPERACIONAL', 'FINALIZAÇÃO', 'NÚMERO DO PROTOCOLO'}
    if not required_columns.issubset(df.columns):
        return "As colunas necessárias não foram encontradas no DataFrame."

    # Remove linhas com valores NaN na coluna 'TEMPO MÉDIO OPERACIONAL'
    df = df.dropna(subset=['TEMPO MÉDIO OPERACIONAL'])

    # Verifica se os valores da coluna 'TEMPO MÉDIO OPERACIONAL' são do tipo timedelta
    if not pd.api.types.is_timedelta64_dtype(df['TEMPO MÉDIO OPERACIONAL']):
        return "A coluna 'TEMPO MÉDIO OPERACIONAL' contém valores que não são do tipo timedelta."

    # Remove duplicatas baseadas no número do protocolo para os casos 'Fora do Escopo'
    df_unique = df.drop_duplicates(subset=['NÚMERO DO PROTOCOLO'])

    # Filtrar apenas tarefas "CADASTRADO" e "ATUALIZADO" (exceto Distribuição)
    df_tmo = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO']) & (df['FILA'] != 'Distribuição')]

    # Agrupar por fila para calcular as quantidades e médias separadas
    tmo_por_carteira = df_tmo.groupby('FILA').agg(
        Quantidade=('FILA', 'size'),
        Cadastrado=('FINALIZAÇÃO', lambda x: (x == 'CADASTRADO').sum()),
        Atualizado=('FINALIZAÇÃO', lambda x: (x == 'ATUALIZADO').sum()),
    ).reset_index()

    # Calcular TMO apenas para "CADASTRADO"
    df_cadastro = df[df['FINALIZAÇÃO'] == 'CADASTRADO'].groupby('FILA')['TEMPO MÉDIO OPERACIONAL'].mean().reset_index()
    df_cadastro.rename(columns={'TEMPO MÉDIO OPERACIONAL': 'TMO Cadastro'}, inplace=True)

    # Calcular TMO apenas para "ATUALIZADO"
    df_atualizacao = df[df['FINALIZAÇÃO'] == 'ATUALIZADO'].groupby('FILA')['TEMPO MÉDIO OPERACIONAL'].mean().reset_index()
    df_atualizacao.rename(columns={'TEMPO MÉDIO OPERACIONAL': 'TMO Atualização'}, inplace=True)

    # Calcular TMO para a fila "Distribuição" considerando apenas "REALIZADO"
    df_distribuicao = df[(df['FILA'] == 'Distribuição') & (df['FINALIZAÇÃO'] == 'REALIZADO')]

    if not df_distribuicao.empty:
        tmo_distribuicao = df_distribuicao.groupby('FILA').agg(
            Quantidade=('FILA', 'size'),
            TMO_Distribuicao=('TEMPO MÉDIO OPERACIONAL', 'mean')
        ).reset_index()
        tmo_distribuicao.rename(columns={'TMO_Distribuicao': 'TMO Cadastro'}, inplace=True)
        tmo_distribuicao['TMO Atualização'] = None  # A fila "Distribuição" não tem TMO Atualização
    else:
        tmo_distribuicao = pd.DataFrame(columns=['FILA', 'Quantidade', 'TMO Cadastro', 'TMO Atualização'])

    # Unir os dados ao DataFrame principal
    tmo_por_carteira = tmo_por_carteira.merge(df_cadastro, on='FILA', how='left')
    tmo_por_carteira = tmo_por_carteira.merge(df_atualizacao, on='FILA', how='left')

    # Adicionar a fila "Distribuição" ao resultado final
    tmo_por_carteira = pd.concat([tmo_por_carteira, tmo_distribuicao], ignore_index=True)

    # Calcular 'Fora do Escopo' considerando apenas protocolos únicos
    fora_do_escopo_contagem = df_unique.groupby('FILA').apply(
        lambda x: x.shape[0] - (x['FINALIZAÇÃO'] == 'CADASTRADO').sum() - (x['FINALIZAÇÃO'] == 'ATUALIZADO').sum()
    ).reset_index(name='Fora do Escopo')

    # Mesclar a contagem de 'Fora do Escopo'
    tmo_por_carteira = tmo_por_carteira.merge(fora_do_escopo_contagem, on='FILA', how='left')

    # Converter os tempos médios para HH:MM:SS
    def format_timedelta(td):
        if pd.isna(td):
            return "00:00:00"
        total_seconds = int(td.total_seconds())
        return f"{total_seconds // 3600:02d}:{(total_seconds % 3600) // 60:02d}:{total_seconds % 60:02d}"

    tmo_por_carteira['TMO Cadastro'] = tmo_por_carteira['TMO Cadastro'].apply(format_timedelta)
    tmo_por_carteira['TMO Atualização'] = tmo_por_carteira['TMO Atualização'].apply(format_timedelta)

    # Selecionar apenas as colunas finais
    tmo_por_carteira = tmo_por_carteira[['FILA', 'Quantidade', 'Cadastrado', 'Atualizado', 'Fora do Escopo', 'TMO Cadastro', 'TMO Atualização']]

    return tmo_por_carteira

def calcular_producao_agrupada(df):
    required_columns = {'FILA', 'FINALIZAÇÃO', 'NÚMERO DO PROTOCOLO'}
    if not required_columns.issubset(df.columns):
        return "As colunas necessárias ('FILA', 'FINALIZAÇÃO', 'NÚMERO DO PROTOCOLO') não foram encontradas no DataFrame."

    df_unique = df.drop_duplicates(subset=['NÚMERO DO PROTOCOLO'])

    grupos = {
        'CAPTURA ANTECIPADA': [' CADASTRO ROBÔ', 'INCIDENTE PROCESSUAL', 'CADASTRO ANS'],
        'SHAREPOINT': ['CADASTRO SHAREPOINT', 'ATUALIZAÇÃO - SHAREPOINT'],
        'CITAÇÃO ELETRÔNICA': ['CADASTRO CITAÇÃO ELETRÔNICA', 'ATUALIZAÇÃO CITAÇÃO ELETRÔNICA'],
        'E-MAIL': ['CADASTRO E-MAIL', 'OFICIOS E-MAIL', 'CADASTRO DE ÓRGÃOS E OFÍCIOS'],
        'PRE CADASTRO E DIJUR': ['PRE CADASTRO E DIJUR']
    }

    df['GRUPO'] = df['FILA'].map(lambda x: next((k for k, v in grupos.items() if x in v), 'OUTROS'))

    df_agrupado = df.groupby('GRUPO').agg(
        Cadastrado=('FINALIZAÇÃO', lambda x: (x == 'CADASTRADO').sum()),
        Atualizado=('FINALIZAÇÃO', lambda x: (x == 'ATUALIZADO').sum()),
        Fora_do_Escopo=('FINALIZAÇÃO', lambda x: ((x != 'CADASTRADO') & (x != 'ATUALIZADO')).sum())
    ).reset_index()

    return df_agrupado

def calcular_producao_email_detalhada(df):
    required_columns = {'FILA', 'FINALIZAÇÃO', 'NÚMERO DO PROTOCOLO', 'TAREFA'}
    if not required_columns.issubset(df.columns):
        return "As colunas necessárias ('FILA', 'FINALIZAÇÃO', 'NÚMERO DO PROTOCOLO', 'TAREFA') não foram encontradas no DataFrame."

    # Filtrando apenas as filas do grupo E-MAIL
    df_email = df[df['FILA'].isin(['CADASTRO E-MAIL', 'OFICIOS', 'CADASTRO DE ÓRGÃOS E OFÍCIOS'])]

    # Separando os de CADASTRO E-MAIL para agrupar por TAREFA
    df_cadastro_email = df_email[df_email['FILA'] == 'CADASTRO E-MAIL']
    df_outros_email = df_email[df_email['FILA'].isin(['OFICIOS', 'CADASTRO DE ÓRGÃOS E OFÍCIOS'])]

    # Agrupando CADASTRO E-MAIL por TAREFA
    df_cadastro_email_agrupado = df_cadastro_email.groupby('TAREFA').agg(
        Quantidade=('FILA', 'size'),
        Cadastrado=('FINALIZAÇÃO', lambda x: (x == 'CADASTRADO').sum()),
        Atualizado=('FINALIZAÇÃO', lambda x: (x == 'ATUALIZADO').sum()),
        Fora_do_Escopo=('FINALIZAÇÃO', lambda x: ((x != 'CADASTRADO') & (x != 'ATUALIZADO')).sum())
    ).reset_index()

    # Agrupando os demais (OFICIOS E-MAIL e CADASTRO DE ÓRGÃOS E OFÍCIOS) por FILA
    df_outros_email_agrupado = df_outros_email.groupby('FILA').agg(
        Quantidade=('FILA', 'size'),
        Cadastrado=('FINALIZAÇÃO', lambda x: (x == 'CADASTRADO').sum()),
        Atualizado=('FINALIZAÇÃO', lambda x: (x == 'ATUALIZADO').sum()),
        Fora_do_Escopo=('FINALIZAÇÃO', lambda x: ((x != 'CADASTRADO') & (x != 'ATUALIZADO')).sum())
    ).reset_index().rename(columns={'FILA': 'TAREFA'})

    # Concatenando os resultados
    df_email_final = pd.concat([df_cadastro_email_agrupado, df_outros_email_agrupado], ignore_index=True)

    return df_email_final

def calcular_e_exibir_tmo_cadastro_atualizacao_por_fila (df_analista, format_timedelta_hms, st):
    """
    Calcula e exibe o TMO médio de Cadastro e Atualização por Fila,
    junto com a quantidade de tarefas realizadas, na dashboard do Streamlit.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - format_timedelta_hms: Função para formatar a duração do TMO em HH:MM:SS.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """
    if 'FILA' in df_analista.columns and 'FINALIZAÇÃO' in df_analista.columns:
        # Filtrar apenas as tarefas finalizadas com CADASTRADO e ATUALIZADO
        filas_finalizadas_analista = df_analista[df_analista['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO'])]

        # Agrupar os dados por 'FILA' e calcular a quantidade de tarefas por fila
        df_quantidade = filas_finalizadas_analista.groupby('FILA').size().reset_index(name='Quantidade')

        # Calcular o TMO médio para cada fila separadamente
        df_tmo_cadastro = filas_finalizadas_analista[filas_finalizadas_analista['FINALIZAÇÃO'] == 'CADASTRADO'].groupby('FILA')['TEMPO MÉDIO OPERACIONAL'].mean().reset_index()
        df_tmo_atualizacao = filas_finalizadas_analista[filas_finalizadas_analista['FINALIZAÇÃO'] == 'ATUALIZADO'].groupby('FILA')['TEMPO MÉDIO OPERACIONAL'].mean().reset_index()

        # Renomear colunas
        df_tmo_cadastro.rename(columns={'TEMPO MÉDIO OPERACIONAL': 'TMO_Cadastro'}, inplace=True)
        df_tmo_atualizacao.rename(columns={'TEMPO MÉDIO OPERACIONAL': 'TMO_Atualizacao'}, inplace=True)

        # Unir os DataFrames pela Fila
        df_resultado = df_quantidade.merge(df_tmo_cadastro, on='FILA', how='left').merge(df_tmo_atualizacao, on='FILA', how='left')

        # Substituir valores NaN por Timedelta(0)
        df_resultado.fillna(pd.Timedelta(seconds=0), inplace=True)

        # Converter os TMOs para HH:MM:SS
        df_resultado['TMO_Cadastro'] = df_resultado['TMO_Cadastro'].apply(format_timedelta_hms)
        df_resultado['TMO_Atualizacao'] = df_resultado['TMO_Atualizacao'].apply(format_timedelta_hms)

        # Renomear colunas para exibição
        df_resultado.rename(columns={
            'FILA': 'Fila',
            'Quantidade': 'Quantidade',
            'TMO_Cadastro': 'TMO Cadastro',
            'TMO_Atualizacao': 'TMO Atualização'
        }, inplace=True)

        # Estilizar tabela no Streamlit
        styled_df = df_resultado.style.format({'Quantidade': '{:.0f}', 'TMO Cadastro': '{}', 'TMO Atualização': '{}'}).set_properties(**{'text-align': 'left'})
        styled_df = styled_df.set_table_styles([dict(selector='th', props=[('text-align', 'left')])])

        # Exibir DataFrame estilizado na dashboard
        st.dataframe(styled_df, hide_index=True, use_container_width=True)
    else:
        st.warning("As colunas necessárias ('FILA' e 'FINALIZAÇÃO') não foram encontradas no DataFrame.")

def calcular_e_exibir_tmo_por_fila(df_analista, analista_selecionado, format_timedelta, st):
    """
    Calcula e exibe o TMO médio por fila, junto com a quantidade de tarefas realizadas, 
    para um analista específico, na dashboard Streamlit.

    Parâmetros:
        - df_analista: DataFrame contendo os dados de análise.
        - analista_selecionado: Nome do analista selecionado.
        - format_timedelta: Função para formatar a duração do TMO em minutos e segundos.
        - st: Referência para o módulo Streamlit (necessário para exibir os resultados).
    """
    if 'FILA' in df_analista.columns:
        # Filtrar apenas as tarefas finalizadas para cálculo do TMO
        filas_finalizadas_analista = df_analista[df_analista['SITUAÇÃO DA TAREFA'] == 'Finalizada']
        
        # Agrupa por 'FILA' e calcula a quantidade e o TMO médio para cada fila
        carteiras_analista = filas_finalizadas_analista.groupby('FILA').agg(
            Quantidade=('FILA', 'size'),
            TMO_médio=('TEMPO MÉDIO OPERACIONAL', 'mean')
        ).reset_index()

        # Converte o TMO médio para minutos e segundos
        carteiras_analista['TMO_médio'] = carteiras_analista['TMO_médio'].apply(format_timedelta)

        # Renomeia as colunas
        carteiras_analista = carteiras_analista.rename(columns={
            'FILA': 'Fila', 
            'Quantidade': 'Quantidade', 
            'TMO_médio': 'TMO Médio por Fila'
        })
        
        # Configura o estilo do DataFrame para alinhar o conteúdo à esquerda
        styled_df = carteiras_analista.style.format({'Quantidade': '{:.0f}', 'TMO Médio por Fila': '{:s}'}).set_properties(**{'text-align': 'left'})
        styled_df = styled_df.set_table_styles([dict(selector='th', props=[('text-align', 'left')])])

        # Exibe a tabela com as colunas Fila, Quantidade e TMO Médio
        st.dataframe(styled_df, hide_index=True, use_container_width=True)
    else:
        st.write("A coluna 'FILA' não foi encontrada no dataframe.")
        carteiras_analista = pd.DataFrame({'Fila': [], 'Quantidade': [], 'TMO Médio por Fila': []})
        styled_df = carteiras_analista.style.format({'Quantidade': '{:.0f}', 'TMO Médio por Fila': '{:s}'}).set_properties(**{'text-align': 'left'})
        styled_df = styled_df.set_table_styles([dict(selector='th', props=[('text-align', 'left')])])
        st.dataframe(styled_df, hide_index=True, use_container_width=True)

def calcular_tmo_por_mes(df):
    # Converter coluna de tempo para timedelta se necessário
    if df['TEMPO MÉDIO OPERACIONAL'].dtype != 'timedelta64[ns]':
        df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')

    # Converter coluna de data corretamente
    df['DATA DE CONCLUSÃO DA TAREFA'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA'], errors='coerce')

    # Remover registros sem data válida
    df = df[df['DATA DE CONCLUSÃO DA TAREFA'].notna()]

    # Extrair AnoMes no formato de período (ano-mês)
    df['AnoMes'] = df['DATA DE CONCLUSÃO DA TAREFA'].dt.to_period('M')

    # Filtrar protocolos finalizados
    df_finalizados = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO', 'REALIZADO'])]

    # Agrupar e calcular somatório e média
    df_tmo_mes = df_finalizados.groupby('AnoMes').agg(
        Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),
        Total_Protocolos=('TEMPO MÉDIO OPERACIONAL', 'count')
    ).reset_index()

    # Calcular TMO médio em minutos
    df_tmo_mes['TMO'] = (df_tmo_mes['Tempo_Total'] / pd.Timedelta(minutes=1)) / df_tmo_mes['Total_Protocolos']

    # Formatar AnoMes como "Abril de 2024"
    df_tmo_mes['AnoMes'] = df_tmo_mes['AnoMes'].dt.to_timestamp().dt.strftime('%B de %Y').str.capitalize()

    return df_tmo_mes[['AnoMes', 'TMO']]

# Função de formatação
def format_timedelta_mes(minutes):
    """Formata um valor em minutos (float) como 'Xh Ym Zs' se acima de 60 minutos, caso contrário, 'X min Ys'."""
    if minutes >= 60:
        hours = int(minutes // 60)
        minutes_remainder = int(minutes % 60)
        seconds = (minutes - hours * 60 - minutes_remainder) * 60
        seconds_int = round(seconds)
        return f"{hours}h {minutes_remainder}m {seconds_int}s"
    else:
        minutes_int = int(minutes)
        seconds = (minutes - minutes_int) * 60
        seconds_int = round(seconds)
        return f"{minutes_int} min {seconds_int}s"

def exibir_tmo_por_mes(df):
    """
    Exibe um gráfico de barras agrupadas do TMO mensal (Geral, Cadastro, Atualização, Auditoria).
    """
    df = df.copy()

    # Converter para timedelta se necessário
    if not pd.api.types.is_timedelta64_dtype(df['TEMPO MÉDIO OPERACIONAL']):
        df['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df['TEMPO MÉDIO OPERACIONAL'], errors='coerce')

    # Adicionar coluna AnoMes
    df['AnoMes'] = pd.to_datetime(df['DATA DE CONCLUSÃO DA TAREFA'], errors='coerce').dt.to_period('M').astype(str)

    # Separar por tipo
    df_geral = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO', 'REALIZADO'])]
    df_cadastro = df[df['FINALIZAÇÃO'] == 'CADASTRADO']
    df_atualizacao = df[df['FINALIZAÇÃO'] == 'ATUALIZADO']
    df_auditoria = df[df['FINALIZAÇÃO'] == 'AUDITADO']

    def calcular_tmo(df_base, nome_col):
        df_base = df_base.copy()
        if df_base.empty:
            return pd.DataFrame(columns=['AnoMes', nome_col])
        agrupado = df_base.groupby('AnoMes')['TEMPO MÉDIO OPERACIONAL'].mean().reset_index()
        agrupado.columns = ['AnoMes', nome_col]
        return agrupado

    df_tmo_geral = calcular_tmo(df_geral, 'TMO_Geral')
    df_tmo_cadastro = calcular_tmo(df_cadastro, 'TMO_Cadastro')
    df_tmo_atualizacao = calcular_tmo(df_atualizacao, 'TMO_Atualizacao')
    df_tmo_auditoria = calcular_tmo(df_auditoria, 'TMO_Auditoria')

    df_tmo_final = df_tmo_geral.merge(df_tmo_cadastro, on='AnoMes', how='left')
    df_tmo_final = df_tmo_final.merge(df_tmo_atualizacao, on='AnoMes', how='left')
    df_tmo_final = df_tmo_final.merge(df_tmo_auditoria, on='AnoMes', how='left')

    # Formatar tempos
    for col in ['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria']:
        df_tmo_final[col + '_Formatado'] = df_tmo_final[col].apply(
            lambda td: f"{int(td.total_seconds() // 3600):02}:{int((td.total_seconds() % 3600) // 60):02}:{int(td.total_seconds() % 60):02}" if pd.notnull(td) else "00:00:00"
        )

    # Filtro de meses
    meses_disponiveis = df_tmo_final['AnoMes'].unique()
    meses_selecionados = st.multiselect("Selecione os meses para exibição", options=meses_disponiveis, default=meses_disponiveis)
    df_tmo_filtrado = df_tmo_final[df_tmo_final['AnoMes'].isin(meses_selecionados)]

    if df_tmo_filtrado.empty:
        st.warning("Nenhum dado disponível para os meses selecionados.")
        return

    # Formato longo para gráfico
    df_long = df_tmo_filtrado.melt(
        id_vars=['AnoMes'],
        value_vars=['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria'],
        var_name='Tipo de TMO',
        value_name='Tempo Médio Operacional'
    )

    format_dict = df_tmo_filtrado.set_index('AnoMes')[
        ['TMO_Geral_Formatado', 'TMO_Cadastro_Formatado', 'TMO_Atualizacao_Formatado', 'TMO_Auditoria_Formatado']
    ].stack().reset_index()
    format_dict.columns = ['AnoMes', 'Tipo de TMO', 'Tempo Formatado']
    format_dict['Tipo de TMO'] = format_dict['Tipo de TMO'].str.replace('_Formatado', '')
    format_map = format_dict.set_index(['AnoMes', 'Tipo de TMO'])['Tempo Formatado'].to_dict()

    cores = {
        'TMO_Geral': '#ff6a1c',
        'TMO_Cadastro': '#d1491c',
        'TMO_Atualizacao': '#a3330f',
        'TMO_Auditoria': '#4b0082'
    }

    labels_legenda = {
        'TMO_Geral': 'Geral',
        'TMO_Cadastro': 'Cadastro',
        'TMO_Atualizacao': 'Atualização',
        'TMO_Auditoria': 'Auditoria'
    }

    df_long['Texto_Rotulo'] = df_long.apply(
        lambda row: f"{labels_legenda[row['Tipo de TMO']]} - {format_map.get((row['AnoMes'], row['Tipo de TMO']), '')}",
        axis=1
    )

    fig = px.bar(
        df_long,
        x='AnoMes',
        y='Tempo Médio Operacional',
        color='Tipo de TMO',
        text='Texto_Rotulo',
        barmode='group',
        labels={'AnoMes': 'Mês', 'Tempo Médio Operacional': 'Tempo Médio Operacional (HH:MM:SS)'},
        color_discrete_map=cores
    )

    fig.update_layout(bargap=0.2, bargroupgap=0.15, showlegend=False)
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

    df_formatado = df_tmo_filtrado[[
        'AnoMes', 'TMO_Geral_Formatado', 'TMO_Cadastro_Formatado', 'TMO_Atualizacao_Formatado', 'TMO_Auditoria_Formatado']]
    df_formatado.columns = ['Mês', 'TMO Geral', 'TMO Cadastro', 'TMO Atualização', 'TMO Auditoria']
    st.dataframe(df_formatado, use_container_width=True, hide_index=True)
        
def exibir_dataframe_tmo_formatado(df):
    # Calcule o TMO mensal usando a função `calcular_tmo_por_mes`
    df_tmo_mes = calcular_tmo_por_mes(df)
    
    # Verifique se há dados para exibir
    if df_tmo_mes.empty:
        st.warning("Nenhum dado finalizado disponível para calcular o TMO mensal.")
        return None
    
    # Adicionar a coluna "Tempo Médio Operacional" com base no TMO calculado
    df_tmo_mes['Tempo Médio Operacional'] = df_tmo_mes['TMO'].apply(format_timedelta_mes)
    df_tmo_mes['Mês'] = df_tmo_mes['AnoMes']
    
    # Selecionar as colunas para exibição
    df_tmo_formatado = df_tmo_mes[['Mês', 'Tempo Médio Operacional']]
    
    st.dataframe(df_tmo_formatado, use_container_width=True, hide_index=True)
    
    return df_tmo_formatado

def export_dataframe(df):
    st.subheader("Exportar Dados")
    
    # Seleção de colunas
    colunas_disponiveis = list(df.columns)
    colunas_selecionadas = st.multiselect(
        "Selecione as colunas que deseja exportar:", colunas_disponiveis, default=[]
    )
    
    # Filtrar o DataFrame pelas colunas selecionadas
    if colunas_selecionadas:
        df_filtrado = df[colunas_selecionadas]
        
        # Botão de download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_filtrado.to_excel(writer, index=False, sheet_name='Dados_Exportados')
        buffer.seek(0)
        
        st.download_button(
            label="Baixar Excel",
            data=buffer,
            file_name="dados_exportados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Selecione pelo menos uma coluna para exportar.")
        
def calcular_melhor_tmo_por_dia(df_analista):
    """
    Calcula o melhor TMO de cadastro por dia para o analista.

    Parâmetros:
        - df_analista: DataFrame filtrado para o analista.

    Retorna:
        - O dia com o melhor TMO de cadastro e o valor do TMO.
    """
    # Filtrar apenas as finalizações do tipo 'CADASTRADO'
    df_cadastro = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']

    # Calcula o TMO por dia para o tipo 'CADASTRADO'
    df_tmo_por_dia = calcular_tmo_por_dia(df_cadastro)

    # Identifica o dia com o menor TMO
    if not df_tmo_por_dia.empty:
        melhor_dia = df_tmo_por_dia.loc[df_tmo_por_dia['TMO'].idxmin()]
        return melhor_dia['Dia'], melhor_dia['TMO']

    # Retorna None caso não haja dados para 'CADASTRADO'
    return None, None

def calcular_melhor_dia_por_cadastro(df_analista):
    # Agrupa os dados por dia e conta os cadastros
    if 'FINALIZAÇÃO' in df_analista.columns and 'DATA DE CONCLUSÃO DA TAREFA' in df_analista.columns:
        df_cadastros_por_dia = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO'].groupby(
            df_analista['DATA DE CONCLUSÃO DA TAREFA'].dt.date
        ).size().reset_index(name='Quantidade')
        
        # Identifica o dia com maior quantidade de cadastros
        if not df_cadastros_por_dia.empty:
            melhor_dia = df_cadastros_por_dia.loc[df_cadastros_por_dia['Quantidade'].idxmax()]
            return melhor_dia['DATA DE CONCLUSÃO DA TAREFA'], melhor_dia['Quantidade']
    
    return None, 0

def exibir_tmo_por_mes_analista(df_analista, analista_selecionado):
    """
    Exibe o gráfico e a tabela do TMO mensal para um analista específico com filtro por mês.

    Parâmetros:
        - df_analista: DataFrame filtrado para o analista.
        - analista_selecionado: Nome do analista selecionado.
    """
    # Calcular o TMO por mês
    df_tmo_mes = calcular_tmo_por_mes(df_analista)

    # Verificar se há dados para exibir
    colunas_existentes = [col for col in ['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria'] if col in df_tmo_mes.columns]
    if not colunas_existentes or df_tmo_mes[colunas_existentes].isna().all(axis=None):
        st.warning(f"Não há dados suficientes para calcular o TMO mensal do analista {analista_selecionado}.")
        return None

    # Formatar o TMO para exibição
    df_tmo_mes['TMO_Formatado'] = df_tmo_mes['TMO'].apply(format_timedelta_mes)

    # Criar multiselect para os meses disponíveis
    meses_disponiveis = df_tmo_mes['AnoMes'].unique()
    meses_selecionados = st.multiselect(
        "Selecione os meses para exibição",
        options=meses_disponiveis,
        default=meses_disponiveis
    )

    # Filtrar os dados com base nos meses selecionados
    df_tmo_mes_filtrado = df_tmo_mes[df_tmo_mes['AnoMes'].isin(meses_selecionados)]

    # Verificar se há dados após o filtro
    if df_tmo_mes_filtrado.empty:
        st.warning("Nenhum dado disponível para os meses selecionados.")
        return None

    # Criar e exibir o gráfico de barras
    fig = px.bar(
        df_tmo_mes_filtrado,
        x='AnoMes',
        y='TMO',
        labels={'AnoMes': 'Mês', 'TMO': 'TMO (minutos)'},
        text=df_tmo_mes_filtrado['TMO_Formatado'],  # Usar o TMO formatado como rótulo
        color_discrete_sequence=['#ff571c', '#7f2b0e', '#4c1908', '#ff884d', '#a34b28', '#331309']
    )
    fig.update_xaxes(type='category')  # Tratar o eixo X como categórico
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

    # Criar e exibir a tabela com os dados formatados
    df_tmo_mes_filtrado['Mês'] = df_tmo_mes_filtrado['AnoMes']  # Renomear para exibição
    df_tmo_formatado = df_tmo_mes_filtrado[['Mês', 'TMO_Formatado']].rename(columns={'TMO_Formatado': 'Tempo Médio Operacional'})
    st.dataframe(df_tmo_formatado, use_container_width=True, hide_index=True)

    return df_tmo_formatado


def calcular_grafico_tmo_analista_por_mes(df_analista):
    """
    Calcula o TMO Geral, Cadastro, Atualização e Auditoria por mês para um analista específico.
    
    Parâmetro:
        - df_analista: DataFrame contendo as tarefas do analista.
    
    Retorna:
        - DataFrame com TMO_Geral, TMO_Cadastro, TMO_Atualizacao e TMO_Auditoria por mês.
    """
    if df_analista.empty:
        return pd.DataFrame(columns=['AnoMes', 'TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria'])
    
    # Garantir tipo timedelta
    if not pd.api.types.is_timedelta64_dtype(df_analista['TEMPO MÉDIO OPERACIONAL']):
        df_analista['TEMPO MÉDIO OPERACIONAL'] = pd.to_timedelta(df_analista['TEMPO MÉDIO OPERACIONAL'], errors='coerce')

    df_analista['AnoMes'] = df_analista['DATA DE CONCLUSÃO DA TAREFA'].dt.to_period('M').astype(str)

    # Separar os subconjuntos
    df_geral = df_analista[df_analista['FINALIZAÇÃO'].isin(['CADASTRADO', 'ATUALIZADO', 'REALIZADO', 'AUDITADO'])]
    df_cadastro = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']
    df_atualizacao = df_analista[df_analista['FINALIZAÇÃO'] == 'ATUALIZADO']
    df_auditoria = df_analista[df_analista['FINALIZAÇÃO'] == 'AUDITADO']

    # Função de cálculo por mês
    def calcular_tmo(df, nome_coluna):
        if df.empty:
            return pd.DataFrame(columns=['AnoMes', nome_coluna])
        df_tmo = df.groupby('AnoMes').agg(
            Tempo_Total=('TEMPO MÉDIO OPERACIONAL', 'sum'),
            Total_Protocolos=('TEMPO MÉDIO OPERACIONAL', 'count')
        ).reset_index()
        df_tmo[nome_coluna] = df_tmo['Tempo_Total'] / df_tmo['Total_Protocolos']
        return df_tmo[['AnoMes', nome_coluna]]

    # Calcular os TMO por tipo
    df_tmo_geral = calcular_tmo(df_geral, 'TMO_Geral')
    df_tmo_cadastro = calcular_tmo(df_cadastro, 'TMO_Cadastro')
    df_tmo_atualizacao = calcular_tmo(df_atualizacao, 'TMO_Atualizacao')
    df_tmo_auditoria = calcular_tmo(df_auditoria, 'TMO_Auditoria')

    # Mesclar todos os resultados
    df_tmo_mes = (
        df_tmo_geral
        .merge(df_tmo_cadastro, on='AnoMes', how='left')
        .merge(df_tmo_atualizacao, on='AnoMes', how='left')
        .merge(df_tmo_auditoria, on='AnoMes', how='left')
    )

    # Preencher valores ausentes e formatar a coluna de mês
    df_tmo_mes.fillna(pd.Timedelta(seconds=0), inplace=True)
    df_tmo_mes['AnoMes'] = pd.to_datetime(df_tmo_mes['AnoMes'], errors='coerce')
    df_tmo_mes['AnoMes'] = df_tmo_mes['AnoMes'].dt.strftime('%B de %Y').str.capitalize()

    return df_tmo_mes

def format_timedelta_grafico_tmo_analista(td):
    """Formata um timedelta no formato HH:MM:SS"""
    if pd.isna(td) or td == pd.Timedelta(seconds=0):
        return "00:00:00"
    
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    seconds = total_seconds % 60

    return f"{hours:02}:{minutes:02}:{seconds:02}"

def format_timedelta_hms(timedelta_value):
    """Formata um timedelta em HH:MM:SS"""
    total_seconds = int(timedelta_value.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

def exibir_grafico_tmo_analista_por_mes(df_analista, analista_selecionado):
    """
    Exibe um gráfico de barras agrupadas do TMO mensal (Geral, Cadastro, Atualização, Auditoria) para um analista específico.
    """

    # Calcular o TMO por mês (a função abaixo precisa calcular TMO_Auditoria também)
    df_tmo_mes = calcular_grafico_tmo_analista_por_mes(df_analista)

    if df_tmo_mes[['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria']].isna().all(axis=None):
        st.warning(f"Não há dados suficientes para calcular o TMO mensal do analista {analista_selecionado}.")
        return None

    # Formatar os tempos para HH:MM:SS
    for col in ['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria']:
        df_tmo_mes[col + '_Formatado'] = df_tmo_mes[col].apply(format_timedelta_hms)

    # Multiselect para meses
    meses_disponiveis = df_tmo_mes['AnoMes'].unique()
    meses_selecionados = st.multiselect(
        "Selecione os meses para exibição",
        options=meses_disponiveis,
        default=meses_disponiveis
    )

    df_tmo_mes_filtrado = df_tmo_mes[df_tmo_mes['AnoMes'].isin(meses_selecionados)]

    if df_tmo_mes_filtrado.empty:
        st.warning("Nenhum dado disponível para os meses selecionados.")
        return None

    # Transformar em formato longo
    df_tmo_long = df_tmo_mes_filtrado.melt(
        id_vars=['AnoMes'], 
        value_vars=['TMO_Geral', 'TMO_Cadastro', 'TMO_Atualizacao', 'TMO_Auditoria'], 
        var_name='Tipo de TMO', 
        value_name='Tempo Médio Operacional'
    )

    # Mapeamento de tempos formatados
    format_dict = df_tmo_mes_filtrado.set_index('AnoMes')[
        ['TMO_Geral_Formatado', 'TMO_Cadastro_Formatado', 'TMO_Atualizacao_Formatado', 'TMO_Auditoria_Formatado']
    ].stack().reset_index()
    format_dict.columns = ['AnoMes', 'Tipo de TMO', 'Tempo Formatado']
    format_dict['Tipo de TMO'] = format_dict['Tipo de TMO'].str.replace('_Formatado', '')
    format_map = format_dict.set_index(['AnoMes', 'Tipo de TMO'])['Tempo Formatado'].to_dict()

    # Cores
    custom_colors = {
        'TMO_Geral': '#ff6a1c',
        'TMO_Cadastro': '#d1491c',
        'TMO_Atualizacao': '#a3330f',
        'TMO_Auditoria': '#4b0082'
    }

    tipo_tmo_label = {
        'TMO_Geral': 'Geral',
        'TMO_Cadastro': 'Cadastro',
        'TMO_Atualizacao': 'Atualização',
        'TMO_Auditoria': 'Auditoria'
    }

    df_tmo_long['Texto_Rotulo'] = df_tmo_long.apply(
        lambda row: f"{tipo_tmo_label[row['Tipo de TMO']]} - {format_map.get((row['AnoMes'], row['Tipo de TMO']), '')}",
        axis=1
    )

    # Gráfico
    fig = px.bar(
        df_tmo_long,
        x='AnoMes',
        y='Tempo Médio Operacional',
        color='Tipo de TMO',
        text='Texto_Rotulo',
        barmode='group',
        labels={'AnoMes': 'Mês', 'Tempo Médio Operacional': 'Tempo Médio Operacional (HH:MM:SS)'},
        color_discrete_map=custom_colors
    )

    fig.update_layout(bargap=0.2, bargroupgap=0.15, showlegend=False)
    fig.update_traces(textposition='outside')
    st.plotly_chart(fig, use_container_width=True)

    # Tabela
    df_tmo_formatado = df_tmo_mes_filtrado[[
        'AnoMes', 'TMO_Geral_Formatado', 'TMO_Cadastro_Formatado', 'TMO_Atualizacao_Formatado', 'TMO_Auditoria_Formatado'
    ]]
    df_tmo_formatado.columns = ['Mês', 'TMO Geral', 'TMO Cadastro', 'TMO Atualização', 'TMO Auditoria']
    st.dataframe(df_tmo_formatado, use_container_width=True, hide_index=True)

def exportar_planilha_com_tmo(df, periodo_selecionado, analistas_selecionados, tmo_tipo='GERAL'):
    """
    Exporta uma planilha com informações do período selecionado, analistas, TMO (geral, cadastrado ou cadastrado com tipo) e quantidade de tarefas,
    adicionando formatação condicional baseada na média do TMO.

    Parâmetros:
        - df: DataFrame com os dados.
        - periodo_selecionado: Tuple contendo a data inicial e final.
        - analistas_selecionados: Lista de analistas selecionados.
        - tmo_tipo: Tipo de TMO a ser usado ('GERAL', 'CADASTRADO', 'CADASTRADO_DETALHADO').
    """
    # Filtrar o DataFrame com base no período e analistas selecionados
    data_inicial, data_final = periodo_selecionado
    df_filtrado = df[
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) &
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final) &
        (df['USUÁRIO QUE CONCLUIU A TAREFA'].isin(analistas_selecionados))
    ]

    # Calcular o TMO e a quantidade por analista
    analistas = []
    tmos = []
    quantidades = []
    tipos_causa = []  # Para armazenar os tipos de "TP CAUSA (TP COMPLEMENTO)"

    for analista in analistas_selecionados:
        df_analista = df_filtrado[df_filtrado['USUÁRIO QUE CONCLUIU A TAREFA'] == analista]
        if tmo_tipo == 'GERAL':
            # Considerar apenas as finalizações "CADASTRADO", "REALIZADO" e "ATUALIZADO"
            df_relevante = df_analista[df_analista['FINALIZAÇÃO'].isin(['CADASTRADO', 'REALIZADO', 'ATUALIZADO'])]
        elif tmo_tipo == 'CADASTRADO':
            # Considerar apenas as finalizações "CADASTRADO"
            df_relevante = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']
        elif tmo_tipo == 'CADASTRADO_DETALHADO':
            # Considerar apenas as finalizações "CADASTRADO" e detalhar por "TP CAUSA (TP COMPLEMENTO)"
            df_relevante = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']
            causa_detalhes = df_relevante.groupby('TP CAUSA (TP COMPLEMENTO)').size().reset_index(name='Quantidade')
            tipos_causa.append(causa_detalhes)
        else:
            st.error("Tipo de TMO inválido selecionado.")
            return

        tmo_analista = calcular_tmo_personalizado(df_relevante)
        quantidade_analista = len(df_relevante)

        analistas.append(analista)
        tmos.append(tmo_analista)
        quantidades.append(quantidade_analista)

    # Criar o DataFrame de resumo
    df_resumo = pd.DataFrame({
        'Analista': analistas,
        'TMO': tmos,
        'Quantidade': quantidades
    })

    # Adicionar o período ao DataFrame exportado
    df_resumo['Período Inicial'] = data_inicial
    df_resumo['Período Final'] = data_final

    # Formatar o TMO como HH:MM:SS
    df_resumo['TMO'] = df_resumo['TMO'].apply(
        lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}"
    )

    # Calcular a média do TMO em segundos
    tmo_segundos = [timedelta(hours=int(t.split(":")[0]), minutes=int(t.split(":")[1]), seconds=int(t.split(":")[2])).total_seconds() for t in df_resumo['TMO']]
    media_tmo_segundos = sum(tmo_segundos) / len(tmo_segundos)

    # Criar um arquivo Excel em memória
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Exportar os dados do resumo
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo')

        # Se for CADASTRADO_DETALHADO, incluir os tipos de causa
        if tmo_tipo == 'CADASTRADO_DETALHADO' and tipos_causa:
            for i, causa_detalhes in enumerate(tipos_causa):
                causa_detalhes.to_excel(writer, index=False, sheet_name=f'Tipos_{analistas[i]}')

        # Acessar o workbook e worksheet para aplicar formatação condicional
        workbook = writer.book
        worksheet = writer.sheets['Resumo']

        # Ajustar largura das colunas
        worksheet.set_column('A:A', 20)  # Coluna 'Analista'
        worksheet.set_column('B:B', 12)  # Coluna 'TMO'
        worksheet.set_column('C:C', 15)  # Coluna 'Quantidade'
        worksheet.set_column('D:E', 15)  # Colunas 'Período Inicial' e 'Período Final'

        # Formatação baseada na média do TMO
        format_tmo_green = workbook.add_format({'bg_color': '#CCFFCC', 'font_color': '#006600'})  # Verde
        format_tmo_yellow = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': '#666600'})  # Amarelo
        format_tmo_red = workbook.add_format({'bg_color': '#FFCCCC', 'font_color': '#FF0000'})  # Vermelho

        # Aplicar formatação condicional
        for row, tmo in enumerate(tmo_segundos, start=2):
            if tmo < media_tmo_segundos * 0.9:  # Abaixo da média
                worksheet.write(f'B{row}', df_resumo.loc[row-2, 'TMO'], format_tmo_green)
            elif media_tmo_segundos * 0.9 <= tmo <= media_tmo_segundos * 1.1:  # Na média ou próximo
                worksheet.write(f'B{row}', df_resumo.loc[row-2, 'TMO'], format_tmo_yellow)
            else:  # Acima da média
                worksheet.write(f'B{row}', df_resumo.loc[row-2, 'TMO'], format_tmo_red)

    buffer.seek(0)

    # Oferecer download
    st.download_button(
        label="Baixar Planilha",
        data=buffer,
        file_name="resumo_analistas_formatado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import timedelta

def exportar_planilha_com_tmo_completo(df, periodo_selecionado, analistas_selecionados):
    """
    Exporta uma planilha com informações do período selecionado, incluindo:
    - TMO de Cadastro
    - Quantidade de Cadastro
    - TMO de Atualizado
    - Quantidade de Atualização
    """

    # Filtrar o DataFrame com base no período e analistas selecionados
    data_inicial, data_final = periodo_selecionado
    df_filtrado = df[
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) &
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final) &
        (df['USUÁRIO QUE CONCLUIU A TAREFA'].isin(analistas_selecionados))
    ]

    # Criar listas para armazenar os dados por analista
    analistas = []
    tmo_cadastrado = []
    quantidade_cadastrado = []
    tmo_atualizado = []
    quantidade_atualizado = []

    for analista in analistas_selecionados:
        df_analista = df_filtrado[df_filtrado['USUÁRIO QUE CONCLUIU A TAREFA'] == analista]

        # Cálculo para Cadastro
        df_cadastro = df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO']
        tmo_cadastro_analista = df_cadastro['TEMPO MÉDIO OPERACIONAL'].mean()
        total_cadastro = len(df_cadastro)

        # Cálculo para Atualizado
        df_atualizado = df_analista[df_analista['FINALIZAÇÃO'] == 'ATUALIZADO']
        tmo_atualizado_analista = df_atualizado['TEMPO MÉDIO OPERACIONAL'].mean()
        total_atualizado = len(df_atualizado)

        # Adicionar valores às listas
        analistas.append(analista)
        tmo_cadastrado.append(tmo_cadastro_analista)
        quantidade_cadastrado.append(total_cadastro)
        tmo_atualizado.append(tmo_atualizado_analista)
        quantidade_atualizado.append(total_atualizado)

    # Criar DataFrame de resumo
    df_resumo = pd.DataFrame({
        'Analista': analistas,
        'TMO Cadastro': tmo_cadastrado,
        'Quantidade Cadastro': quantidade_cadastrado,
        'TMO Atualizado': tmo_atualizado,
        'Quantidade Atualização': quantidade_atualizado
    })

    # Converter TMO para HH:MM:SS (removendo frações de segundos)
    def format_tmo(tmo):
        if pd.notnull(tmo):
            total_seconds = int(tmo.total_seconds())  # Removendo frações
            return str(timedelta(seconds=total_seconds))
        return '00:00:00'

    df_resumo['TMO Cadastro'] = df_resumo['TMO Cadastro'].apply(format_tmo)
    df_resumo['TMO Atualizado'] = df_resumo['TMO Atualizado'].apply(format_tmo)

    # Criar um arquivo Excel em memória
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df_resumo.to_excel(writer, index=False, sheet_name='Resumo')

        # Ajustes no Excel
        workbook = writer.book
        worksheet = writer.sheets['Resumo']
        worksheet.set_column('A:A', 20)  # Coluna Analista
        worksheet.set_column('B:E', 15)  # Colunas de TMO e Quantidade

    buffer.seek(0)

    # Oferecer download no Streamlit
    st.download_button(
        label="Baixar Planilha Completa de TMO",
        data=buffer,
        file_name="relatorio_tmo_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def exportar_relatorio_detalhado_por_analista(df, periodo_selecionado, analistas_selecionados):
    """
    Exporta um relatório detalhado por analista, com TMO de CADASTRADO e quantidade por dia, gerando uma aba para cada analista.

    Parâmetros:
        - df: DataFrame com os dados.
        - periodo_selecionado: Tuple contendo a data inicial e final.
        - analistas_selecionados: Lista de analistas selecionados.
    """
    data_inicial, data_final = periodo_selecionado

    # Filtrar o DataFrame pelo período e analistas selecionados
    df_filtrado = df[
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) &
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final) &
        (df['USUÁRIO QUE CONCLUIU A TAREFA'].isin(analistas_selecionados)) &
        (df['FINALIZAÇÃO'] == 'CADASTRADO')  # Apenas tarefas cadastradas
    ]

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        # Criar relatório detalhado por analista
        for analista in analistas_selecionados:
            df_analista = df_filtrado[df_filtrado['USUÁRIO QUE CONCLUIU A TAREFA'] == analista]
            
            # Calcular TMO e quantidade por dia
            df_tmo_por_dia = df_analista.groupby(df_analista['DATA DE CONCLUSÃO DA TAREFA'].dt.date).agg(
                TMO=('TEMPO MÉDIO OPERACIONAL', lambda x: x.sum() / len(x) if len(x) > 0 else pd.Timedelta(0)),
                Quantidade=('DATA DE CONCLUSÃO DA TAREFA', 'count')
            ).reset_index()
            df_tmo_por_dia.rename(columns={'DATA DE CONCLUSÃO DA TAREFA': 'Dia'}, inplace=True)

            # Formatar TMO como HH:MM:SS
            df_tmo_por_dia['TMO'] = df_tmo_por_dia['TMO'].apply(
                lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}"
            )

            # Adicionar coluna de analista
            df_tmo_por_dia.insert(0, 'Analista', analista)

            # Reordenar colunas para "ANALISTA, TMO, QUANTIDADE, DIA"
            df_tmo_por_dia = df_tmo_por_dia[['Analista', 'TMO', 'Quantidade', 'Dia']]

            # Exportar os dados para uma aba do Excel
            if not df_tmo_por_dia.empty:
                df_tmo_por_dia.to_excel(writer, index=False, sheet_name=analista[:31])

                # Acessar a aba para formatação condicional
                workbook = writer.book
                worksheet = writer.sheets[analista[:31]]

                # Ajustar largura das colunas
                worksheet.set_column('A:A', 20)  # Coluna 'Analista'
                worksheet.set_column('B:B', 12)  # Coluna 'TMO'
                worksheet.set_column('C:C', 12)  # Coluna 'Quantidade'
                worksheet.set_column('D:D', 15)  # Coluna 'Dia'

                # Criar formatos para formatação condicional
                format_tmo_green = workbook.add_format({'bg_color': '#CCFFCC', 'font_color': '#006600'})  # Verde
                format_tmo_yellow = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': '#666600'})  # Amarelo
                format_tmo_red = workbook.add_format({'bg_color': '#FFCCCC', 'font_color': '#FF0000'})  # Vermelho

                # Aplicar formatação condicional com base no TMO
                worksheet.conditional_format(
                    'B2:B{}'.format(len(df_tmo_por_dia) + 1),
                    {
                        'type': 'formula',
                        'criteria': f'=LEN(B2)>0',
                        'format': format_tmo_yellow  # Formato padrão para demonstração
                    }
                )

    buffer.seek(0)

    # Oferecer download
    st.download_button(
        label="Baixar Relatório Detalhado por Analista (TMO CADASTRADO)",
        data=buffer,
        file_name="relatorio_tmo_detalhado_cadastrado_por_analista.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def calcular_tmo_geral(df):
    """
    Calcula o TMO Geral considerando todas as tarefas finalizadas.
    """
    df_finalizados = df[df['FINALIZAÇÃO'].isin(['CADASTRADO', 'REALIZADO', 'ATUALIZADO'])]
    return df_finalizados['TEMPO MÉDIO OPERACIONAL'].mean()

def calcular_tmo_cadastro(df):
    """
    Calcula o TMO apenas para tarefas finalizadas como "CADASTRADO".
    """
    df_cadastro = df[df['FINALIZAÇÃO'] == 'CADASTRADO']
    return df_cadastro['TEMPO MÉDIO OPERACIONAL'].mean()

def calcular_tempo_ocioso(df):
    """
    Calcula o tempo ocioso total por analista.
    """
    df_ocioso = df.groupby('USUÁRIO QUE CONCLUIU A TAREFA')['TEMPO OCIOSO'].sum().reset_index()
    return df_ocioso

def gerar_relatorio_tmo_completo(df, periodo_selecionado, analistas_selecionados):
    """
    Gera um relatório Excel com TMO de Cadastro, TMO Geral, Quantidade de Cadastro,
    Quantidade Total de Protocolos e Tempo Ocioso.
    """
    data_inicial, data_final = periodo_selecionado

    # 🔹 Filtrar os dados pelo período e analistas selecionados
    df_filtrado = df[
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicial) &
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_final) &
        (df['USUÁRIO QUE CONCLUIU A TAREFA'].isin(analistas_selecionados))
    ].copy()  # Criar uma cópia para evitar alterações no DataFrame original

    # 🔹 Calcular o tempo ocioso por analista
    df_tempo_ocioso = calcular_tempo_ocioso_por_analista(df_filtrado)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        for analista in analistas_selecionados:
            df_analista = df_filtrado[df_filtrado['USUÁRIO QUE CONCLUIU A TAREFA'] == analista]
            tmo_geral = calcular_tmo_geral(df_analista)
            tmo_cadastro = calcular_tmo_cadastro(df_analista)
            total_cadastros = len(df_analista[df_analista['FINALIZAÇÃO'] == 'CADASTRADO'])
            total_protocolos = len(df_analista)

            # 🔹 Ajuste para acessar a coluna correta do DataFrame `df_tempo_ocioso`
            tempo_ocioso = df_tempo_ocioso[df_tempo_ocioso['USUÁRIO QUE CONCLUIU A TAREFA'] == analista]['TEMPO OCIOSO'].sum() if not df_tempo_ocioso.empty else pd.Timedelta(0)

            # 🔹 Criar DataFrame com os dados do relatório
            df_resumo = pd.DataFrame({
                'Analista': [analista],
                'TMO Geral': [tmo_geral],
                'TMO Cadastro': [tmo_cadastro],
                'Quantidade Cadastro': [total_cadastros],
                'Quantidade Total': [total_protocolos],
                'Tempo Ocioso': [tempo_ocioso]
            })

            # 🔹 Converter TMO e Tempo Ocioso para HH:MM:SS
            df_resumo['TMO Geral'] = df_resumo['TMO Geral'].apply(lambda x: str(timedelta(seconds=x.total_seconds())) if pd.notnull(x) else '00:00:00')
            df_resumo['TMO Cadastro'] = df_resumo['TMO Cadastro'].apply(lambda x: str(timedelta(seconds=x.total_seconds())) if pd.notnull(x) else '00:00:00')
            df_resumo['Tempo Ocioso'] = df_resumo['Tempo Ocioso'].apply(lambda x: str(timedelta(seconds=x.total_seconds())) if pd.notnull(x) else '00:00:00')

            # 🔹 Escrever no Excel
            df_resumo.to_excel(writer, index=False, sheet_name=analista[:31])

            # 🔹 Ajustes no Excel
            workbook = writer.book
            worksheet = writer.sheets[analista[:31]]
            worksheet.set_column('A:A', 20)
            worksheet.set_column('B:F', 15)

    buffer.seek(0)

    # 🔹 Botão de download no Streamlit
    st.download_button(
        label="📊 Baixar Relatório Completo de TMO",
        data=buffer,
        file_name="relatorio_tmo_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

import json

def gerar_relatorio_html(df, data_inicio_antes, data_fim_antes, data_inicio_depois, data_fim_depois, usuarios_selecionados):
    df = df[df['USUÁRIO QUE CONCLUIU A TAREFA'].isin(usuarios_selecionados)]

    def filtrar_periodo(df, inicio, fim):
        return df[(df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= inicio) &
                  (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= fim)]

    df_antes = filtrar_periodo(df, data_inicio_antes, data_fim_antes)
    df_depois = filtrar_periodo(df, data_inicio_depois, data_fim_depois)

    def calcular_tmo_por_tipo(df_periodo, tipo):
        return df_periodo[df_periodo['FINALIZAÇÃO'] == tipo].groupby('USUÁRIO QUE CONCLUIU A TAREFA')['TEMPO MÉDIO OPERACIONAL'].mean()

    def format_tmo(value):
        if pd.isnull(value) or value == pd.Timedelta(0):
            return "00:00:00"
        total_seconds = value.total_seconds()
        return f"{int(total_seconds // 3600):02}:{int((total_seconds % 3600) // 60):02}:{int(total_seconds % 60):02}"

    analistas = sorted(set(df_antes['USUÁRIO QUE CONCLUIU A TAREFA']).union(df_depois['USUÁRIO QUE CONCLUIU A TAREFA']))

    tabela_html = ""
    tmo_antes_list = []
    tmo_depois_list = []
    tmo_antes_legenda = []
    tmo_depois_legenda = []
    nomes_analistas = []

    for analista in analistas:
        tmo_antes_cadastro = calcular_tmo_por_tipo(df_antes, 'CADASTRADO').get(analista, pd.Timedelta(0))
        tmo_depois_cadastro = calcular_tmo_por_tipo(df_depois, 'CADASTRADO').get(analista, pd.Timedelta(0))
        tmo_antes_atualizacao = calcular_tmo_por_tipo(df_antes, 'ATUALIZADO').get(analista, pd.Timedelta(0))
        tmo_depois_atualizacao = calcular_tmo_por_tipo(df_depois, 'ATUALIZADO').get(analista, pd.Timedelta(0))

        tabela_html += f"""
        <tr>
            <td>{analista}</td>
            <td>{format_tmo(tmo_antes_cadastro)}</td>
            <td>{format_tmo(tmo_depois_cadastro)}</td>
            <td>{format_tmo(tmo_antes_atualizacao)}</td>
            <td>{format_tmo(tmo_depois_atualizacao)}</td>
        </tr>
        """

        nomes_analistas.append(analista)
        tmo_antes_list.append(int(tmo_antes_cadastro.total_seconds() // 60))
        tmo_depois_list.append(int(tmo_depois_cadastro.total_seconds() // 60))
        tmo_antes_legenda.append(format_tmo(tmo_antes_cadastro))
        tmo_depois_legenda.append(format_tmo(tmo_depois_cadastro))

    label_antes = f"TMO ({data_inicio_antes.strftime('%d/%m')} - {data_fim_antes.strftime('%d/%m')})"
    label_depois = f"TMO ({data_inicio_depois.strftime('%d/%m')} - {data_fim_depois.strftime('%d/%m')})"

    tmo_antes_legenda_json = json.dumps(tmo_antes_legenda)
    tmo_depois_legenda_json = json.dumps(tmo_depois_legenda)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Relatório de TMO</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 0; padding: 20px; }}
            .container {{ max-width: 1000px; background-color: white; padding: 20px; border-radius: 10px; margin: auto; }}
            .header {{ text-align: center; padding-bottom: 20px; }}
            .header h1 {{ color: #333; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; text-align: center; padding: 10px; }}
            th {{ background-color: #FF5500; color: white; }}
            .header img {{ width: 150px; margin: 10px auto; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="https://finchsolucoes.com.br/img/fefdd9df-1bd3-4107-ab22-f06d392c1f55.png" alt="Finch Soluções">
                <h1>Relatório de TMO</h1>
                <h2>Comparação entre Períodos</h2>
            </div>
            <canvas id="tmoChart" width="400" height="200"></canvas>
            <script>
                Chart.register(ChartDataLabels);
                const tmoAntesLegenda = {tmo_antes_legenda_json};
                const tmoDepoisLegenda = {tmo_depois_legenda_json};
                var ctx = document.getElementById('tmoChart').getContext('2d');
                var tmoChart = new Chart(ctx, {{
                    type: 'bar',
                    data: {{
                        labels: {json.dumps(nomes_analistas)},
                        datasets: [
                            {{
                                label: '{label_antes}',
                                data: {tmo_antes_list},
                                backgroundColor: '#FF5500',
                                borderRadius: 10
                            }},
                            {{
                                label: '{label_depois}',
                                data: {tmo_depois_list},
                                backgroundColor: '#330066',
                                borderRadius: 10
                            }}
                        ]
                    }},
                    options: {{
                        responsive: true,
                        plugins: {{
                            datalabels: {{
                                anchor: 'end',
                                align: 'top',
                                color: '#000',
                                font: {{ size: 10 }},
                                formatter: function(value, context) {{
                                    return context.dataset.label === '{label_antes}' 
                                        ? tmoAntesLegenda[context.dataIndex] 
                                        : tmoDepoisLegenda[context.dataIndex];
                                }}
                            }}
                        }},
                        scales: {{
                            y: {{
                                beginAtZero: true
                            }}
                        }}
                    }}
                }});
            </script>
            <table>
                <tr>
                    <th>Analista</th>
                    <th>TMO Cadastro ({data_inicio_antes.strftime('%d/%m')} - {data_fim_antes.strftime('%d/%m')})</th>
                    <th>TMO Cadastro ({data_inicio_depois.strftime('%d/%m')} - {data_fim_depois.strftime('%d/%m')})</th>
                    <th>TMO Atualização ({data_inicio_antes.strftime('%d/%m')} - {data_fim_antes.strftime('%d/%m')})</th>
                    <th>TMO Atualização ({data_inicio_depois.strftime('%d/%m')} - {data_fim_depois.strftime('%d/%m')})</th>
                </tr>
                {tabela_html}
            </table>
        </div>
    </body>
    </html>
    """

    return html_content

# **🔹 Função para baixar o HTML**
def download_html(df, data_inicio_antes, data_fim_antes, data_inicio_depois, data_fim_depois, usuarios_selecionados):
    html_content = gerar_relatorio_html(df, data_inicio_antes, data_fim_antes, data_inicio_depois, data_fim_depois, usuarios_selecionados)
    buffer = BytesIO()
    buffer.write(html_content.encode("utf-8"))
    buffer.seek(0)

    st.download_button(
        label="Baixar Relatório em HTML",
        data=buffer,
        file_name="relatorio_tmo.html",
        mime="text/html"
    )

def gerar_relatorio_html_tmo(df, data_inicio, data_fim):
    """
    Gera um relatório HTML de TMO de Cadastro com um gráfico e tabela detalhada.

    Parâmetros:
        - df: DataFrame contendo os dados de TMO.
        - data_inicio, data_fim: Período selecionado.
    """

    # Filtrar apenas cadastros e o período selecionado
    df_filtrado = df[
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date >= data_inicio) &
        (df['DATA DE CONCLUSÃO DA TAREFA'].dt.date <= data_fim) &
        (df['FINALIZAÇÃO'] == 'CADASTRADO')
    ]

    # Calcular TMO médio geral
    tmo_medio_geral = df_filtrado['TEMPO MÉDIO OPERACIONAL'].mean()
    
    # Agrupar dados por analista
    df_tmo_analista = df_filtrado.groupby('USUÁRIO QUE CONCLUIU A TAREFA').agg(
        TMO=('TEMPO MÉDIO OPERACIONAL', lambda x: x.mean() if len(x) > 0 else pd.Timedelta(0)),
        Quantidade=('DATA DE CONCLUSÃO DA TAREFA', 'count')
    ).reset_index()

    # Formatar TMO para exibição
    df_tmo_analista['TMO'] = df_tmo_analista['TMO'].apply(
        lambda x: f"{int(x.total_seconds() // 3600):02}:{int((x.total_seconds() % 3600) // 60):02}:{int(x.total_seconds() % 60):02}"
    )

    # Organizar os dados para gráfico
    nomes_analistas = df_tmo_analista['USUÁRIO QUE CONCLUIU A TAREFA'].tolist()
    tmo_valores = [int(pd.Timedelta(tmo).total_seconds() // 60) for tmo in df_tmo_analista['TMO']]  # Converter para minutos
    tmo_labels = df_tmo_analista['TMO'].tolist()

    # Criar a tabela HTML
    tabela_html = ""
    for _, row in df_tmo_analista.iterrows():
        tabela_html += f"""
        <tr>
            <td>{row['USUÁRIO QUE CONCLUIU A TAREFA']}</td>
            <td>{row['TMO']}</td>
            <td>{row['Quantidade']}</td>
            <td>{data_inicio.strftime('%B/%Y')}</td>
        </tr>
        """

    # Criar o HTML final
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Relatório de Produtividade</title>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels"></script>
        <style>
            body {{ font-family: Arial, sans-serif; background-color: #f4f4f4; margin: 0; padding: 20px; }}
            .container {{ max-width: 1280px; background-color: white; padding: 20px; border-radius: 15px; margin: auto; }}
            .header {{ text-align: center; padding-bottom: 20px; }}
            .header h1 {{ color: #333; }}
            .highlight {{ background-color: #FF5500; color: white; padding: 10px; text-align: center; border-radius: 10px; font-size: 16px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            table, th, td {{ border: 1px solid #ddd; text-align: center; }}
            th {{ background-color: #FF5500; color: white; padding: 10px; }}
            td {{ padding: 10px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <img src="https://finchsolucoes.com.br/img/fefdd9df-1bd3-4107-ab22-f06d392c1f55.png" alt="Finch Soluções" width="150px">
                <h1>Relatório de Produtividade</h1>
                <h2>{data_inicio.strftime('%B/%Y')}</h2>
            </div>
            <div class="highlight">
                <p style="font-size: 25px;"><strong>Média de TMO de Cadastro</strong> <br>{tmo_medio_geral}</p>
            </div>

            <canvas id="tmoChart" width="400" height="200"></canvas>

            <script>
                Chart.register(ChartDataLabels);
                var ctx = document.getElementById('tmoChart').getContext('2d');
                var tmoData = {tmo_labels};
                var tmoChart = new Chart(ctx, {{
                    type: 'bar',
                    data: {{
                        labels: {nomes_analistas},
                        datasets: [{{
                            label: 'TMO de Cadastro',
                            data: {tmo_valores},
                            backgroundColor: ['#330066', '#FF5500', '#330066', '#FF5500', '#330066', '#FF5500', '#330066', '#FF5500', '#330066'],
                            borderRadius: 10
                        }}]
                    }},
                    options: {{
                        responsive: true,
                        plugins: {{
                            datalabels: {{
                                anchor: 'end',
                                align: 'top',
                                formatter: (value, ctx) => tmoData[ctx.dataIndex],
                                color: '#000'
                            }},
                            tooltip: {{
                                callbacks: {{
                                    label: function(tooltipItem) {{
                                        return tmoData[tooltipItem.dataIndex];
                                    }}
                                }}
                            }}
                        }},
                        scales: {{
                            y: {{
                                beginAtZero: true
                            }}
                        }}
                    }}
                }});
            </script>

            <table>
                <tr>
                    <th>Analista</th>
                    <th>TMO de Cadastro</th>
                    <th>Quantidade de Cadastro</th>
                    <th>Mês de Referência</th>
                </tr>
                {tabela_html}
            </table>
        </div>
    </body>
    </html>
    """

    return html_content


def download_html_tmo(df, data_inicio, data_fim):
    """
    Função para gerar e baixar o relatório HTML de TMO.
    """
    html_content = gerar_relatorio_html_tmo(df, data_inicio, data_fim)
    buffer = BytesIO()
    buffer.write(html_content.encode("utf-8"))
    buffer.seek(0)

    st.download_button(
        label="Baixar Relatório HTML de TMO",
        data=buffer,
        file_name="relatorio_tmo.html",
        mime="text/html"
    )

# FILAS - INCIDENTE, CADASTRO ROBO E CADASTRO ANS - CONTAGEM DA QUANTIDADE DE TAREFAS QEU ENTRARAM POR DIA (PANDAS)
# CRIAÇÃO DO PROTOCOLO -> .cont()
# finalização NA - TIRAR A SIUTAÇÃO COMO CANCELADA E VERIFICAR DESTINO DA TAREFA

# --- NOVA FUNÇÃO: Contagem de Desvios na Fila de Auditoria ---
import pandas as pd
import plotly.express as px
import streamlit as st

def contar_desvios(df):
    """
    Conta a frequência de cada tipo de desvio presente na coluna 'DESVIOS CADASTRO'.
    Os desvios podem estar separados por vírgula na mesma célula.

    Retorna:
        DataFrame com colunas ['Desvio', 'Frequência']
    """
    if 'DESVIOS CADASTRO' not in df.columns:
        return pd.DataFrame({'Desvio': [], 'Frequência': []})

    # Remove valores nulos, separa por vírgula e explode
    df_exploded = df['DESVIOS CADASTRO'].dropna().str.split(',').explode().str.strip()

    # Conta a frequência de cada desvio individual
    df_contagem = df_exploded.value_counts().reset_index()
    df_contagem.columns = ['Desvio', 'Frequência']

    return df_contagem

def exibir_grafico_desvios_auditoria(df):
    """
    Exibe um gráfico de barras com a frequência de desvios encontrados na coluna 'DESVIOS CADASTRO'.
    """
    df_desvios = contar_desvios(df)

    if df_desvios.empty:
        st.info("Nenhum desvio encontrado na coluna 'DESVIOS CADASTRO'.")
        return

    fig = px.bar(
        df_desvios,
        x='Desvio',
        y='Frequência',
        text='Frequência',
        color='Desvio',
        color_discrete_sequence=px.colors.sequential.Reds[::-1]
    )

    fig.update_layout(
        xaxis_title="Tipo de Desvio",
        yaxis_title="Frequência",
        xaxis_tickangle=-45,
        showlegend=False
    )
    st.plotly_chart(fig, use_container_width=True)
    
        # Exibir o desvio mais incidente como métrica
    desvio_top = df_desvios.iloc[0]
    with st.container(border=True):
        st.metric(label="Desvio Mais Frequente", value=desvio_top['Desvio'], delta=f"{desvio_top['Frequência']} ocorrências")
