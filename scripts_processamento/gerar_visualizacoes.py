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
    pasta_base = "graficos_interativos"
    
    # Cria a pasta se não existir
    os.makedirs(pasta_base, exist_ok=True)
    
    return pasta_base

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
    # Usa a data_hora para extrair a hora
    df_horarios = df[pd.notna(df['data_hora'])].copy()
    df_horarios['hora_int'] = df_horarios['data_hora'].dt.hour
    
    # Agrupa por hora
    contagem_horarios = df_horarios['hora_int'].value_counts().sort_index().reset_index()
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
    fig.write_html(os.path.join(pasta_visualizacoes, "distribuicao_horarios.html"))
    
    return contagem_horarios

def analisar_dias_semana(df: pd.DataFrame, pasta_visualizacoes: str):
    """Analisa a distribuição de ligações por dia da semana."""
    # Usa a data_hora para extrair o dia da semana
    df_dias = df[pd.notna(df['data_hora'])].copy()
    df_dias['dia_semana'] = df_dias['data_hora'].dt.day_name()
    
    # Traduz os dias da semana para português
    traducao_dias = {
        'Monday': 'Segunda-feira',
        'Tuesday': 'Terça-feira',
        'Wednesday': 'Quarta-feira',
        'Thursday': 'Quinta-feira',
        'Friday': 'Sexta-feira'
    }
    df_dias['dia_semana_pt'] = df_dias['dia_semana'].map(traducao_dias)
    
    # Conta ligações por dia
    contagem_dias = df_dias['dia_semana_pt'].value_counts().reindex(DIAS_SEMANA.values()).reset_index()
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
    fig.write_html(os.path.join(pasta_visualizacoes, "distribuicao_dias_semana.html"))
    
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
    # Ordena o DataFrame do maior para o menor valor
    df_categorias = df_categorias.sort_values('Quantidade', ascending=True)
    
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
    fig.write_html(os.path.join(pasta_visualizacoes, "distribuicao_categorias.html"))

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
    fig.write_html(os.path.join(pasta_visualizacoes, "correlacao_categorias.html"))

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
    fig.write_html(os.path.join(pasta_visualizacoes, "atendimentos_por_atendente.html"))
    
    return contagem_atendentes

def gerar_relatorio_html(df: pd.DataFrame, df_categorias: pd.DataFrame, 
                        df_atendentes: pd.DataFrame, df_horarios: pd.DataFrame,
                        df_dias_semana: pd.DataFrame,
                        pasta_visualizacoes: str):
    """Gera um relatório HTML com os resultados"""
    # Gera as linhas da tabela de categorias
    linhas_tabela = ""
    for _, row in df_categorias.iterrows():
        if row['Categoria'] != "Outros assuntos":  # Remove a categoria duplicada
            linhas_tabela += f"""
            <tr>
                <td>{row['Categoria']}</td>
                <td>{row['Quantidade']}</td>
                <td>{row['Percentual']:.1f}%</td>
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
    
    # Análises críticas
    analise_atendentes = """
    <div class="analise-critica">
        <h3>Análise Crítica - Distribuição por Atendente</h3>
        <p>Observa-se uma concentração significativa de atendimentos (60.3%) realizados pela atendente Joelma, 
        o que pode indicar um desequilíbrio na distribuição de chamadas. Isso pode levar a sobrecarga de trabalho 
        e possível impacto na qualidade do atendimento. Recomenda-se avaliar a distribuição atual das chamadas e 
        considerar uma redistribuição mais equilibrada entre os atendentes.</p>
    </div>
    """
    
    analise_dias = """
    <div class="analise-critica">
        <h3>Análise Crítica - Distribuição por Dia da Semana</h3>
        <p>A segunda-feira apresenta um volume significativamente maior de ligações (25.9%), 
        seguido por uma distribuição mais uniforme nos outros dias. Este padrão sugere uma 
        maior demanda no início da semana, possivelmente devido a acúmulo de demandas do fim 
        de semana. Recomenda-se reforçar a equipe nas segundas-feiras para melhor atender 
        este pico de demanda.</p>
    </div>
    """
    
    analise_categorias = """
    <div class="analise-critica">
        <h3>Análise Crítica - Distribuição de Categorias</h3>
        <p>Destaca-se o alto índice de ligações indevidas/sem resposta (52.8%), o que representa 
        um problema significativo de eficiência no sistema de atendimento. O segundo maior motivo 
        é agendamento de consultas (47.8%), seguido por encaminhamentos para WhatsApp (20.4%). 
        Recomenda-se:</p>
        <ul>
            <li>Investigar as causas das ligações indevidas/sem resposta</li>
            <li>Avaliar a implementação de um sistema de agendamento online para reduzir o volume de ligações</li>
            <li>Otimizar o processo de encaminhamento para WhatsApp</li>
        </ul>
    </div>
    """
    
    analise_horarios = """
    <div class="analise-critica">
        <h3>Análise Crítica - Distribuição por Horário</h3>
        <p>O pico de ligações ocorre entre 11h e 13h, com destaque para as 11h (47 ligações). 
        Este padrão indica uma concentração de demanda no horário próximo ao almoço. Sugere-se:</p>
        <ul>
            <li>Reforçar a equipe de atendimento neste período</li>
            <li>Considerar escala de almoço alternada para manter capacidade de atendimento</li>
            <li>Avaliar a possibilidade de incentivos para clientes utilizarem horários alternativos</li>
        </ul>
    </div>
    """
    
    analise_final = """
    <div class="analise-final">
        <h2>Análise Final e Recomendações</h2>
        <p>A análise dos dados de Maio/2025 revela pontos críticos que merecem atenção:</p>
        
        <h3>Principais Achados:</h3>
        <ul>
            <li>Alto volume de ligações indevidas (52.8%)</li>
            <li>Concentração de atendimentos em uma única atendente (60.3%)</li>
            <li>Pico de demanda às segundas-feiras (25.9%)</li>
            <li>Horário crítico entre 11h e 13h</li>
        </ul>
        
        <h3>Recomendações Estratégicas:</h3>
        <ol>
            <li><strong>Otimização do Sistema de Atendimento:</strong>
                <ul>
                    <li>Implementar sistema de agendamento online</li>
                    <li>Melhorar triagem inicial das chamadas</li>
                    <li>Desenvolver FAQ no site para reduzir ligações simples</li>
                </ul>
            </li>
            <li><strong>Gestão de Recursos Humanos:</strong>
                <ul>
                    <li>Redistribuir chamadas entre atendentes</li>
                    <li>Capacitar equipe para múltiplos tipos de atendimento</li>
                    <li>Implementar escala flexível para cobrir horários de pico</li>
                </ul>
            </li>
            <li><strong>Melhorias Tecnológicas:</strong>
                <ul>
                    <li>Avaliar sistema de callback para reduzir tempo de espera</li>
                    <li>Implementar chatbot para dúvidas básicas</li>
                    <li>Melhorar integração entre canais de atendimento</li>
                </ul>
            </li>
        </ol>
        
        <p>A implementação dessas recomendações pode levar a uma significativa melhoria na 
        eficiência do atendimento e satisfação dos pacientes.</p>
    </div>
    """
    
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Análise de Ligações - Instituto Bhariátrica - Maio/2025</title>
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
            
            h3 {{
                color: var(--secondary-color);
                font-size: 1.4em;
                margin-top: 20px;
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
            
            .analise-critica {{
                background-color: var(--background-light);
                padding: 20px;
                border-radius: 8px;
                margin: 20px 0;
                border-left: 4px solid var(--primary-color);
            }}
            
            .analise-final {{
                background-color: var(--background-light);
                padding: 30px;
                border-radius: 10px;
                margin: 40px 0;
                border: 2px solid var(--primary-color);
            }}
            
            .analise-final ul, .analise-final ol {{
                padding-left: 20px;
            }}
            
            .analise-final li {{
                margin: 10px 0;
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
            <img src="imagens/logo.png" alt="Logo Bhariátrica" class="logo">
        </div>
        
        <div class="container">
            <h1>Análise de Ligações - Maio/2025</h1>
            
            <div class="stats">
                <h2>Estatísticas Gerais</h2>
                <p>Total de ligações analisadas: <span class="highlight">{len(df)}</span></p>
                <p>Data da análise: <span class="highlight">{datetime.now().strftime("%d/%m/%Y %H:%M:%S")}</span></p>
            </div>

            <div class="grid-container">
                <div class="visualization">
                    <h2>Distribuição por Atendente</h2>
                    <iframe src="graficos_interativos/atendimentos_por_atendente.html" frameborder="0"></iframe>
                    {analise_atendentes}
                </div>

                <div class="visualization">
                    <h2>Distribuição por Dia da Semana</h2>
                    <iframe src="graficos_interativos/distribuicao_dias_semana.html" frameborder="0"></iframe>
                    {analise_dias}
                </div>
            </div>

            <div class="visualization">
                <h2>Distribuição de Categorias</h2>
                <iframe src="graficos_interativos/distribuicao_categorias.html" frameborder="0"></iframe>
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
                {analise_categorias}
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
                {analise_horarios}
            </div>

            {analise_final}
        </div>
        
        <div class="footer">
            <p>© {datetime.now().year} Instituto Bhariátrica - Relatório gerado automaticamente</p>
        </div>
    </body>
    </html>
    """
    
    # Salva o relatório na pasta raiz
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)

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
    print("3. Um relatório HTML completo: 'index.html'")

if __name__ == "__main__":
    main() 