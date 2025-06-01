import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import seaborn as sns
import os
from datetime import datetime
import json
from typing import Dict, List
import re

# Mapeamento de ramais para nomes (você pode atualizar com os nomes corretos)
MAPEAMENTO_ATENDENTES = {
    'bioc5310': 'Lucas',
    'bioc5311': 'Ana Rafaela',
    'bioc5313': 'Júlio',
    'bioc5315': 'Lucila',
    'bioc5316': 'Bia',
    'bioc5318': 'Joelma',
    'bioc5319': 'Joelma'
}

# Mapeamento de dias da semana em português
DIAS_SEMANA = {
    0: 'Segunda-feira',
    1: 'Terça-feira',
    2: 'Quarta-feira',
    3: 'Quinta-feira',
    4: 'Sexta-feira'
}

def criar_pasta_visualizacoes() -> str:
    """Cria e retorna o caminho da pasta de visualizações"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pasta_base = os.path.join("scripts_processamento", "visualizacoes")
    pasta_visualizacoes = os.path.join(pasta_base, f"analise_{timestamp}")
    
    # Cria as pastas necessárias
    os.makedirs(pasta_visualizacoes, exist_ok=True)
    os.makedirs(os.path.join(pasta_visualizacoes, "graficos_estaticos"), exist_ok=True)
    os.makedirs(os.path.join(pasta_visualizacoes, "graficos_interativos"), exist_ok=True)
    
    return pasta_visualizacoes

def carregar_dados() -> pd.DataFrame:
    """Carrega os dados mais recentes de classificação"""
    pasta_resultados = os.path.join("scripts_processamento", "bd_transcricao", "2025-05")
    
    # Encontra o arquivo de classificação mais recente
    arquivos = [f for f in os.listdir(pasta_resultados) if f.startswith('classificacao_parcial_')]
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo de classificação encontrado!")
    
    arquivo_mais_recente = max(arquivos)
    caminho_arquivo = os.path.join(pasta_resultados, arquivo_mais_recente)
    
    return pd.read_excel(caminho_arquivo)

def extrair_hora_arquivo(nome_arquivo: str) -> int:
    """Extrai a hora do nome do arquivo de áudio."""
    padrao = r"_(\d{2})h\d{2}"
    match = re.search(padrao, nome_arquivo)
    if match:
        return int(match.group(1))
    return -1

def analisar_horarios(df: pd.DataFrame, pasta_visualizacoes: str):
    """Analisa a distribuição de ligações por horário do dia."""
    # Extrai hora do nome do arquivo
    df['hora'] = df['nome_arquivo'].apply(extrair_hora_arquivo)
    
    # Remove registros com hora inválida
    df_horarios = df[df['hora'] != -1].copy()
    
    # Agrupa por hora
    contagem_horarios = df_horarios['hora'].value_counts().sort_index().reset_index()
    contagem_horarios.columns = ['Hora', 'Quantidade']
    
    # Gera gráfico interativo de barras
    fig = px.bar(contagem_horarios,
                 x='Hora',
                 y='Quantidade',
                 title='Distribuição de Ligações por Hora do Dia',
                 text='Quantidade')
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        height=600,
        showlegend=False,
        title_x=0.5,
        title_font_size=20,
        xaxis_title="Hora do Dia",
        yaxis_title="Quantidade de Ligações",
        xaxis=dict(
            tickmode='array',
            ticktext=[f"{h:02d}h" for h in contagem_horarios['Hora']],
            tickvals=contagem_horarios['Hora']
        )
    )
    
    # Salva o gráfico
    fig.write_html(os.path.join(pasta_visualizacoes, "graficos_interativos", "distribuicao_horarios.html"))
    
    return contagem_horarios

def analisar_dias_semana(df: pd.DataFrame, pasta_visualizacoes: str):
    """Analisa a distribuição de ligações por dia da semana."""
    # Converte a data do nome do arquivo para dia da semana
    df['dia_semana'] = pd.to_datetime(df['data_processamento']).dt.dayofweek
    
    # Filtra apenas dias úteis (0-4: segunda a sexta)
    df_dias = df[df['dia_semana'].between(0, 4)].copy()
    
    # Mapeia números para nomes dos dias
    df_dias['dia_semana_nome'] = df_dias['dia_semana'].map(DIAS_SEMANA)
    
    # Conta ligações por dia
    contagem_dias = df_dias['dia_semana_nome'].value_counts().reindex(DIAS_SEMANA.values()).reset_index()
    contagem_dias.columns = ['Dia', 'Quantidade']
    
    # Gera gráfico interativo
    fig = px.bar(contagem_dias,
                 x='Dia',
                 y='Quantidade',
                 title='Distribuição de Ligações por Dia da Semana',
                 text='Quantidade')
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        height=600,
        showlegend=False,
        title_x=0.5,
        title_font_size=20,
        xaxis_title="Dia da Semana",
        yaxis_title="Quantidade de Ligações"
    )
    
    # Salva o gráfico
    fig.write_html(os.path.join(pasta_visualizacoes, "graficos_interativos", "distribuicao_dias_semana.html"))
    
    return contagem_dias

def preparar_dados_categorias(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara os dados para visualização das categorias"""
    # Extrai todas as categorias únicas
    todas_categorias = []
    for cats in df['categorias_detectadas'].dropna():
        todas_categorias.extend([c.strip() for c in cats.split(',')])
    
    # Conta a frequência de cada categoria
    contagem = pd.Series(todas_categorias).value_counts()
    
    # Calcula percentuais
    percentuais = (contagem / len(df)) * 100
    
    # Cria DataFrame com contagens e percentuais
    df_categorias = pd.DataFrame({
        'Categoria': contagem.index,
        'Quantidade': contagem.values,
        'Percentual': percentuais.values
    })
    
    # Ordena por quantidade em ordem decrescente
    return df_categorias.sort_values('Quantidade', ascending=False)

def gerar_grafico_barras_horizontal(df_categorias: pd.DataFrame, pasta_visualizacoes: str):
    """Gera gráfico de barras horizontal interativo"""
    fig = px.bar(df_categorias, 
                 x='Quantidade', 
                 y='Categoria',
                 orientation='h',
                 title='Distribuição de Categorias nas Ligações',
                 text=df_categorias['Percentual'].round(1).astype(str) + '%')
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        height=800,
        showlegend=False,
        title_x=0.5,
        title_font_size=20
    )
    
    # Salva o gráfico interativo
    fig.write_html(os.path.join(pasta_visualizacoes, "graficos_interativos", "distribuicao_categorias.html"))
    
    # Versão estática para relatórios
    plt.figure(figsize=(12, 8))
    sns.barplot(data=df_categorias, x='Quantidade', y='Categoria')
    plt.title('Distribuição de Categorias nas Ligações')
    plt.tight_layout()
    plt.savefig(os.path.join(pasta_visualizacoes, "graficos_estaticos", "distribuicao_categorias.png"))
    plt.close()

def gerar_grafico_correlacao_categorias(df: pd.DataFrame, pasta_visualizacoes: str):
    """Gera matriz de correlação entre categorias"""
    # Cria colunas binárias para cada categoria
    categorias_unicas = []
    for cats in df['categorias_detectadas'].dropna():
        categorias_unicas.extend([c.strip() for c in cats.split(',')])
    categorias_unicas = list(set(categorias_unicas))
    
    # Cria matriz de correlação
    matriz_categorias = pd.DataFrame(columns=categorias_unicas, index=range(len(df)))
    for idx, row in df.iterrows():
        if pd.isna(row['categorias_detectadas']):
            continue
        cats = [c.strip() for c in row['categorias_detectadas'].split(',')]
        for cat in categorias_unicas:
            matriz_categorias.at[idx, cat] = 1 if cat in cats else 0
    
    # Calcula correlação
    correlacao = matriz_categorias.corr()
    
    # Gera heatmap interativo
    fig = go.Figure(data=go.Heatmap(
        z=correlacao,
        x=correlacao.columns,
        y=correlacao.columns,
        text=correlacao.round(2),
        texttemplate='%{text}',
        textfont={"size": 10},
        hoverongaps=False))
    
    fig.update_layout(
        title='Correlação entre Categorias',
        height=800,
        width=800,
        title_x=0.5
    )
    
    # Salva versão interativa
    fig.write_html(os.path.join(pasta_visualizacoes, "graficos_interativos", "correlacao_categorias.html"))
    
    # Versão estática
    plt.figure(figsize=(12, 10))
    sns.heatmap(correlacao, annot=True, cmap='coolwarm', center=0)
    plt.title('Correlação entre Categorias')
    plt.tight_layout()
    plt.savefig(os.path.join(pasta_visualizacoes, "graficos_estaticos", "correlacao_categorias.png"))
    plt.close()

def extrair_ramal(nome_arquivo: str) -> str:
    """Extrai o ramal do nome do arquivo."""
    padrao = r"bioc(\d{4})"
    match = re.search(padrao, nome_arquivo.lower())
    if match:
        ramal_completo = f"bioc{match.group(1)}"
        # Verifica se o ramal está no intervalo válido (5310-5319)
        ramal_numero = int(match.group(1))
        if 5310 <= ramal_numero <= 5319:
            return ramal_completo
    return "Não identificado"

def get_nome_atendente(ramal: str) -> str:
    """Retorna o nome da atendente com base no ramal."""
    return MAPEAMENTO_ATENDENTES.get(ramal, "Atendente não identificada")

def gerar_grafico_atendentes(df: pd.DataFrame, pasta_visualizacoes: str):
    """Gera gráfico de atendimentos por atendente."""
    # Extrai ramais e converte para nomes
    df['ramal'] = df['nome_arquivo'].apply(extrair_ramal)
    df['atendente'] = df['ramal'].apply(get_nome_atendente)
    
    # Conta atendimentos por atendente
    contagem_atendentes = df['atendente'].value_counts().reset_index()
    contagem_atendentes.columns = ['Atendente', 'Quantidade']
    
    # Gera gráfico interativo
    fig = px.bar(contagem_atendentes,
                 x='Atendente',
                 y='Quantidade',
                 title='Distribuição de Atendimentos por Atendente',
                 text='Quantidade')
    
    fig.update_traces(textposition='outside')
    fig.update_layout(
        height=600,
        showlegend=False,
        title_x=0.5,
        title_font_size=20
    )
    
    # Salva o gráfico
    fig.write_html(os.path.join(pasta_visualizacoes, "graficos_interativos", "atendimentos_por_atendente.html"))
    
    return contagem_atendentes

def gerar_relatorio_html(df: pd.DataFrame, df_categorias: pd.DataFrame, 
                        df_atendentes: pd.DataFrame, df_horarios: pd.DataFrame,
                        df_dias_semana: pd.DataFrame,
                        pasta_visualizacoes: str):
    """Gera um relatório HTML com os resultados"""
    # Gera as linhas da tabela de categorias
    linhas_tabela = ""
    for _, row in df_categorias.iterrows():
        linhas_tabela += f"""
        <tr>
            <td>{row['Categoria']}</td>
            <td>{row['Quantidade']}</td>
            <td>{row['Percentual']:.1f}%</td>
        </tr>
        """
    
    # Gera as linhas da tabela de atendentes
    tabela_atendentes = ""
    for _, row in df_atendentes.iterrows():
        tabela_atendentes += f"""
        <tr>
            <td>{row['Atendente']}</td>
            <td>{row['Quantidade']}</td>
        </tr>
        """
    
    # Gera as linhas da tabela de horários
    tabela_horarios = ""
    for _, row in df_horarios.nlargest(5, 'Quantidade').iterrows():
        tabela_horarios += f"""
        <tr>
            <td>{row['Hora']}:00</td>
            <td>{row['Quantidade']}</td>
        </tr>
        """
    
    # Gera as linhas da tabela de dias da semana
    tabela_dias_semana = ""
    for _, row in df_dias_semana.iterrows():
        tabela_dias_semana += f"""
        <tr>
            <td>{row['Dia']}</td>
            <td>{row['Quantidade']}</td>
        </tr>
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
            
            h1 {{
                color: var(--primary-color);
                text-align: center;
                font-size: 2.5em;
                margin-bottom: 40px;
            }}
            
            h2 {{ 
                color: var(--primary-color);
                font-size: 1.8em;
                margin-top: 30px;
            }}
            
            .stats {{ 
                background-color: var(--background-light);
                padding: 25px;
                border-radius: 10px;
                margin: 20px 0;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            
            .visualization {{ 
                background-color: #ffffff;
                padding: 20px;
                border-radius: 10px;
                margin: 30px 0;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            
            table {{ 
                width: 100%;
                border-collapse: collapse;
                margin: 20px 0;
                background-color: #ffffff;
            }}
            
            th, td {{ 
                padding: 15px;
                text-align: left;
                border-bottom: 1px solid #dee2e6;
            }}
            
            th {{ 
                background-color: var(--primary-color);
                color: #ffffff;
            }}
            
            tr:hover {{
                background-color: var(--background-light);
            }}
            
            .stats p {{
                font-size: 1.1em;
                margin: 10px 0;
            }}
            
            .footer {{
                background-color: var(--primary-color);
                color: #ffffff;
                text-align: center;
                padding: 20px;
                margin-top: 50px;
            }}
            
            iframe {{
                border: none;
                width: 100%;
                height: 600px;
                border-radius: 5px;
            }}
            
            .highlight {{
                color: var(--primary-color);
                font-weight: bold;
            }}

            .grid-container {{
                display: grid;
                grid-template-columns: 1fr 1fr;
                gap: 20px;
                margin: 20px 0;
            }}

            @media (max-width: 768px) {{
                .grid-container {{
                    grid-template-columns: 1fr;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <img src="logo.png" alt="Logo Bhariátrica" class="logo">
        </div>
        
        <div class="container">
            <h1>Análise de Ligações - Relatório Detalhado</h1>
            
            <div class="stats">
                <h2>Estatísticas Gerais</h2>
                <p>Total de ligações analisadas: <span class="highlight">{len(df)}</span></p>
                <p>Data da análise: <span class="highlight">{datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</span></p>
            </div>

            <div class="grid-container">
                <div class="visualization">
                    <h2>Distribuição por Atendente</h2>
                    <iframe src="graficos_interativos/atendimentos_por_atendente.html" frameborder="0"></iframe>
                    <table>
                        <tr>
                            <th>Atendente</th>
                            <th>Quantidade</th>
                        </tr>
                        {tabela_atendentes}
                    </table>
                </div>

                <div class="visualization">
                    <h2>Distribuição por Dia da Semana</h2>
                    <iframe src="graficos_interativos/distribuicao_dias_semana.html" frameborder="0"></iframe>
                    <table>
                        <tr>
                            <th>Dia</th>
                            <th>Quantidade</th>
                        </tr>
                        {tabela_dias_semana}
                    </table>
                </div>
            </div>

            <div class="visualization">
                <h2>Distribuição de Categorias</h2>
                <iframe src="graficos_interativos/distribuicao_categorias.html" frameborder="0"></iframe>
            </div>

            <div class="visualization">
                <h2>Correlação entre Categorias</h2>
                <iframe src="graficos_interativos/correlacao_categorias.html" frameborder="0"></iframe>
            </div>

            <div class="visualization">
                <h2>Distribuição por Horário</h2>
                <iframe src="graficos_interativos/distribuicao_horarios.html" frameborder="0"></iframe>
                <table>
                    <tr>
                        <th>Hora</th>
                        <th>Quantidade</th>
                    </tr>
                    {tabela_horarios}
                </table>
            </div>

            <div class="stats">
                <h2>Detalhamento por Categoria</h2>
                <table>
                    <tr>
                        <th>Categoria</th>
                        <th>Quantidade</th>
                        <th>Percentual</th>
                    </tr>
                    {linhas_tabela}
                </table>
            </div>
        </div>
        
        <div class="footer">
            <p>© {datetime.now().year} Instituto Bhariátrica - Relatório gerado automaticamente</p>
        </div>
    </body>
    </html>
    """
    
    # Salva o relatório na pasta raiz
    with open("relatorio_ligacoes.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    
    # Copia os arquivos de gráficos interativos para uma pasta na raiz
    pasta_graficos = "graficos_interativos"
    os.makedirs(pasta_graficos, exist_ok=True)
    
    # Copia os arquivos de gráficos
    import shutil
    for arquivo in ['distribuicao_categorias.html', 'correlacao_categorias.html',
                   'atendimentos_por_atendente.html', 'distribuicao_horarios.html',
                   'distribuicao_dias_semana.html']:
        shutil.copy2(
            os.path.join(pasta_visualizacoes, "graficos_interativos", arquivo),
            os.path.join(pasta_graficos, arquivo)
        )

def main():
    print("Iniciando geração de visualizações...")
    
    # Cria pasta para as visualizações
    pasta_visualizacoes = criar_pasta_visualizacoes()
    print(f"Pasta de visualizações criada: {pasta_visualizacoes}")
    
    # Carrega os dados
    df = carregar_dados()
    print("Dados carregados com sucesso.")
    
    # Prepara dados das categorias
    df_categorias = preparar_dados_categorias(df)
    print("Dados preparados para visualização.")
    
    # Gera visualizações
    print("Gerando visualizações...")
    gerar_grafico_barras_horizontal(df_categorias, pasta_visualizacoes)
    gerar_grafico_correlacao_categorias(df, pasta_visualizacoes)
    
    # Novas análises
    print("Gerando análises adicionais...")
    df_atendentes = gerar_grafico_atendentes(df, pasta_visualizacoes)
    df_horarios = analisar_horarios(df, pasta_visualizacoes)
    df_dias_semana = analisar_dias_semana(df, pasta_visualizacoes)
    
    # Gera relatório HTML
    print("Gerando relatório HTML...")
    gerar_relatorio_html(df, df_categorias, df_atendentes, df_horarios, df_dias_semana, pasta_visualizacoes)
    
    print("\nProcesso concluído!")
    print(f"Todos os arquivos foram salvos em: {pasta_visualizacoes}")
    print("\nVocê encontrará:")
    print("1. Gráficos estáticos na pasta 'graficos_estaticos'")
    print("2. Gráficos interativos na pasta 'graficos_interativos'")
    print("3. Um relatório HTML completo: 'relatorio_ligacoes.html'")

if __name__ == "__main__":
    main() 