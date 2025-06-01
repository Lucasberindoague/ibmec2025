import pandas as pd
import plotly.express as px
from datetime import datetime, time
import holidays
import numpy as np
import os

def carregar_dados(arquivo):
    """Carrega os dados do arquivo CSV."""
    try:
        # Tenta primeiro com UTF-8
        df = pd.read_csv(arquivo, sep=';', encoding='utf-8')
    except UnicodeDecodeError:
        try:
            # Tenta com latin1
            df = pd.read_csv(arquivo, sep=';', encoding='latin1')
        except:
            # Tenta com cp1252 (Windows-1252)
            df = pd.read_csv(arquivo, sep=';', encoding='cp1252')
    
    # Converte a coluna Data/Hora para datetime
    df['data'] = pd.to_datetime(df['Data/Hora'].str.split(' ').str[0], format='%d/%m/%Y')
    df['hora'] = df['Data/Hora'].str.split(' ').str[1]
    
    # Converte a coluna de tempo total para numérico
    df['tempo_total'] = pd.to_numeric(df['Tempo Total'].str.replace(',', '.'), errors='coerce')
    
    return df

def esta_em_horario_comercial(row):
    """Verifica se a ligação está dentro do horário comercial."""
    hora = int(row['hora'].split(':')[0])
    minuto = int(row['hora'].split(':')[1])
    hora_atual = time(hora, minuto)
    
    # Define horários comerciais
    if row['dia_semana'] == 4:  # Sexta-feira
        inicio = time(8, 0)
        fim = time(17, 0)
    else:  # Segunda a quinta
        inicio = time(8, 0)
        fim = time(18, 0)
    
    return inicio <= hora_atual <= fim

def eh_dia_util(data):
    """Verifica se a data é um dia útil."""
    br_holidays = holidays.BR()
    return (data.weekday() < 5) and (data not in br_holidays)

def validar_duracao(row):
    """Valida a duração mínima da ligação baseado na data."""
    data_corte = pd.Timestamp('2024-04-01')
    duracao_minima = 22 if row['data'] >= data_corte else 50
    return row['tempo_total'] >= duracao_minima

def consolidar_status(status):
    """Consolida os status em atendidas e não atendidas."""
    status = str(status).lower()
    if 'atendida' in status:
        return 'Atendida'
    return 'Não Atendida'

def validar_numero(numero):
    """Valida se o número de telefone tem pelo menos 9 dígitos."""
    numero_str = str(numero)
    digitos = ''.join(filter(str.isdigit, numero_str))
    return len(digitos) >= 9

def processar_dados(df):
    """Aplica todos os filtros e processamentos necessários."""
    # Adiciona coluna de dia da semana (0 = Segunda, 6 = Domingo)
    df['dia_semana'] = df['data'].dt.dayofweek
    
    # Aplica filtros
    df_filtrado = df[
        df.apply(esta_em_horario_comercial, axis=1) &  # Horário comercial
        df['data'].apply(eh_dia_util) &  # Dias úteis
        df.apply(validar_duracao, axis=1) &  # Duração mínima
        df['Origem'].apply(validar_numero)  # Número válido
    ].copy()
    
    # Consolida status
    df_filtrado['status_consolidado'] = df_filtrado['Status'].apply(consolidar_status)
    
    return df_filtrado

def gerar_metricas(df):
    """Gera as métricas solicitadas."""
    metricas = {
        'total_ligacoes': len(df),
        'total_atendidas': len(df[df['status_consolidado'] == 'Atendida']),
        'total_nao_atendidas': len(df[df['status_consolidado'] == 'Não Atendida']),
        'taxa_atendimento': (len(df[df['status_consolidado'] == 'Atendida']) / len(df)) * 100 if len(df) > 0 else 0
    }
    
    # Detalhamento de não atendidas
    detalhamento_nao_atendidas = df[df['status_consolidado'] == 'Não Atendida']['Status'].value_counts()
    metricas['detalhamento_nao_atendidas'] = detalhamento_nao_atendidas.to_dict()
    
    return metricas

def gerar_grafico_horas(df, pasta_saida):
    """Gera gráfico de volume por hora."""
    df['hora_cheia'] = df['hora'].str.split(':').str[0].astype(int)
    volume_por_hora = df['hora_cheia'].value_counts().sort_index()
    
    # Cria DataFrame para plotly
    df_plot = pd.DataFrame({
        'hora': [f'{h:02d}h' for h in volume_por_hora.index],
        'quantidade': volume_por_hora.values
    })
    
    fig = px.bar(
        df_plot,
        x='hora',
        y='quantidade',
        title='Volume de Ligações por Hora',
        labels={'quantidade': 'Quantidade de Ligações', 'hora': 'Hora'}
    )
    
    fig.update_layout(height=600, showlegend=False)
    fig.write_html(f"{pasta_saida}/volume_por_hora.html")
    return volume_por_hora.to_dict()

def gerar_grafico_dias_semana(df, pasta_saida):
    """Gera gráfico de volume por dia da semana."""
    mapeamento_dias = {
        0: 'Segunda-feira',
        1: 'Terça-feira',
        2: 'Quarta-feira',
        3: 'Quinta-feira',
        4: 'Sexta-feira'
    }
    
    df['dia_semana_nome'] = df['dia_semana'].map(mapeamento_dias)
    volume_por_dia = df['dia_semana_nome'].value_counts()
    
    # Reordena os dias da semana
    ordem_dias = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira']
    volume_por_dia = volume_por_dia.reindex(ordem_dias)
    
    # Cria DataFrame para plotly
    df_plot = pd.DataFrame({
        'dia': volume_por_dia.index,
        'quantidade': volume_por_dia.values
    })
    
    fig = px.bar(
        df_plot,
        x='dia',
        y='quantidade',
        title='Volume de Ligações por Dia da Semana',
        labels={'quantidade': 'Quantidade de Ligações', 'dia': 'Dia da Semana'}
    )
    
    fig.update_layout(height=600, showlegend=False)
    fig.write_html(f"{pasta_saida}/volume_por_dia.html")
    return volume_por_dia.to_dict()

def gerar_relatorio_html(metricas, volume_hora, volume_dia, pasta_saida):
    """Gera o relatório HTML com os resultados."""
    criterios_html = """
    <div class="criterios">
        <h2>Critérios de Análise</h2>
        <h3>1. Horário Comercial Válido:</h3>
        <ul>
            <li>Segunda a quinta-feira: 08:00 às 18:00</li>
            <li>Sexta-feira: 08:00 às 17:00</li>
        </ul>
        
        <h3>2. Dias Úteis:</h3>
        <ul>
            <li>Apenas dias úteis (segunda a sexta-feira)</li>
            <li>Excluídos feriados nacionais</li>
        </ul>
        
        <h3>3. Duração Mínima das Ligações:</h3>
        <ul>
            <li>A partir de abril/2024: ≥ 22 segundos</li>
            <li>Antes de abril/2024: ≥ 50 segundos</li>
        </ul>
        
        <h3>4. Status das Ligações:</h3>
        <ul>
            <li>Atendidas: ligações com atendimento completo</li>
            <li>Não Atendidas: ligações perdidas, rejeitadas ou ocupadas</li>
        </ul>
        
        <h3>5. Validação de Números:</h3>
        <ul>
            <li>Considerados apenas números com 9 ou mais dígitos</li>
        </ul>
    </div>
    """
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Análise de Ligações - Instituto Bhariátrica</title>
        <meta charset="UTF-8">
        <style>
            :root {{
                --primary-color: #17a2b8;
                --secondary-color: #6c757d;
                --background-light: #f8f9fa;
                --text-dark: #343a40;
            }}
            
            body {{ 
                font-family: Arial, sans-serif;
                margin: 0;
                padding: 0;
                background-color: #ffffff;
                color: var(--text-dark);
            }}
            
            .header {{
                background-color: #ffffff;
                padding: 20px 0;
                border-bottom: 1px solid #eee;
            }}
            
            .logo {{
                max-width: 200px;
                margin: 0 auto;
                display: block;
            }}
            
            .container {{ 
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
            }}
            
            h1, h2 {{ 
                color: var(--primary-color);
                margin-top: 30px;
            }}
            
            h3 {{
                color: var(--secondary-color);
                margin-top: 20px;
            }}
            
            .criterios {{
                background-color: var(--background-light);
                padding: 25px;
                border-radius: 10px;
                margin: 20px 0;
                border-left: 4px solid var(--primary-color);
            }}
            
            .metricas {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
                gap: 20px;
                margin: 30px 0;
            }}
            
            .metrica-card {{
                background-color: var(--background-light);
                padding: 20px;
                border-radius: 10px;
                text-align: center;
            }}
            
            .metrica-valor {{
                font-size: 2em;
                color: var(--primary-color);
                font-weight: bold;
            }}
            
            .visualization {{
                background-color: #ffffff;
                padding: 20px;
                border-radius: 10px;
                margin: 30px 0;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            
            iframe {{
                border: none;
                width: 100%;
                height: 600px;
                border-radius: 5px;
            }}
            
            .footer {{
                background-color: var(--primary-color);
                color: #ffffff;
                text-align: center;
                padding: 20px;
                margin-top: 50px;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <img src="imagens/logo.png" alt="Logo Bhariátrica" class="logo">
        </div>
        
        <div class="container">
            <h1>Análise de Performance do Atendimento Telefônico</h1>
            
            {criterios_html}
            
            <div class="metricas">
                <div class="metrica-card">
                    <h3>Total de Ligações Válidas</h3>
                    <div class="metrica-valor">{metricas['total_ligacoes']}</div>
                </div>
                
                <div class="metrica-card">
                    <h3>Ligações Atendidas</h3>
                    <div class="metrica-valor">{metricas['total_atendidas']}</div>
                </div>
                
                <div class="metrica-card">
                    <h3>Ligações Não Atendidas</h3>
                    <div class="metrica-valor">{metricas['total_nao_atendidas']}</div>
                </div>
                
                <div class="metrica-card">
                    <h3>Taxa de Atendimento</h3>
                    <div class="metrica-valor">{metricas['taxa_atendimento']:.1f}%</div>
                </div>
            </div>
            
            <div class="visualization">
                <h2>Distribuição por Hora do Dia</h2>
                <iframe src="volume_por_hora.html"></iframe>
            </div>
            
            <div class="visualization">
                <h2>Distribuição por Dia da Semana</h2>
                <iframe src="volume_por_dia.html"></iframe>
            </div>
        </div>
        
        <div class="footer">
            <p>© {datetime.now().year} Instituto Bhariátrica - Relatório gerado automaticamente</p>
        </div>
    </body>
    </html>
    """
    
    with open(f"{pasta_saida}/relatorio_performance.html", "w", encoding="utf-8") as f:
        f.write(html_content)

def main():
    # Configurações
    arquivo_entrada = "BD/LIGAÇÕES RECEBIDAS/12.24 até 05.25.csv"
    pasta_saida = "relatorios/performance"
    
    # Cria pasta de saída
    os.makedirs(pasta_saida, exist_ok=True)
    
    # Carrega e processa os dados
    df = carregar_dados(arquivo_entrada)
    df_processado = processar_dados(df)
    
    # Gera métricas e gráficos
    metricas = gerar_metricas(df_processado)
    volume_hora = gerar_grafico_horas(df_processado, pasta_saida)
    volume_dia = gerar_grafico_dias_semana(df_processado, pasta_saida)
    
    # Gera relatório HTML
    gerar_relatorio_html(metricas, volume_hora, volume_dia, pasta_saida)
    
    # Exporta base tratada
    df_processado.to_excel(f"{pasta_saida}/base_tratada.xlsx", index=False)
    
    print("Análise concluída! Os resultados foram salvos em:", pasta_saida)

if __name__ == "__main__":
    main() 