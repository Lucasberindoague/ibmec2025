import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import re
import os

def extrair_horario(nome_arquivo):
    """Extrai o horário do nome do arquivo."""
    padrao = r"_(\d{2})_(\d{2})_(\d{2})_"
    match = re.search(padrao, nome_arquivo)
    if match:
        hora, minuto, segundo = match.groups()
        return f"{hora}:{minuto}:{segundo}"
    return "00:00:00"  # Horário padrão se não encontrar

def criar_grafico_pizza(dados, titulo, arquivo):
    """Cria um gráfico de pizza com os dados fornecidos."""
    plt.figure(figsize=(12, 8))
    cores = sns.color_palette('husl', n_colors=len(dados))
    patches, texts, autotexts = plt.pie(dados.values, labels=dados.index, colors=cores,
                                      autopct='%1.1f%%', startangle=90)
    # Melhorando a legibilidade dos rótulos
    plt.setp(autotexts, size=8, weight="bold")
    plt.setp(texts, size=8)
    plt.title(titulo, pad=20, size=12, weight="bold")
    plt.axis('equal')
    plt.savefig(arquivo, bbox_inches='tight', dpi=300)
    plt.close()

def gerar_relatorio():
    """Gera relatório e visualizações a partir dos dados analisados."""
    # Pasta para salvar os resultados
    pasta_resultados = "resultados"
    pasta_graficos = os.path.join(pasta_resultados, "graficos")
    os.makedirs(pasta_graficos, exist_ok=True)
    
    # Carregar dados analisados
    df = pd.read_csv(os.path.join(pasta_resultados, "analisado.csv"))
    
    # Adicionar coluna de horário
    df['horario'] = df['arquivo'].apply(extrair_horario)
    df['hora'] = df['horario'].str[:2].astype(int)
    
    # Data atual para os arquivos
    data_atual = datetime.now().strftime("%Y%m%d")
    
    # Análise por hora do dia
    plt.figure(figsize=(15, 6))
    contagem_horas = df['hora'].value_counts().sort_index()
    sns.barplot(x=contagem_horas.index, y=contagem_horas.values)
    plt.title("Distribuição de Ligações por Hora do Dia")
    plt.xlabel("Hora")
    plt.ylabel("Quantidade de Ligações")
    plt.xticks(range(24))
    plt.tight_layout()
    plt.savefig(os.path.join(pasta_graficos, f"distribuicao_horaria_{data_atual}.png"))
    plt.close()

    # Gráficos de pizza
    # 1. Distribuição de Categorias
    criar_grafico_pizza(
        df['categoria'].value_counts(),
        'Distribuição de Categorias das Ligações',
        os.path.join(pasta_graficos, f"categorias_pizza_{data_atual}.png")
    )
    
    # 2. Distribuição de Sentimentos
    criar_grafico_pizza(
        df['sentimento'].value_counts(),
        'Distribuição de Sentimentos das Ligações',
        os.path.join(pasta_graficos, f"sentimentos_pizza_{data_atual}.png")
    )

    # Análise de categorias por período do dia
    df['periodo'] = pd.cut(df['hora'], 
                         bins=[0, 12, 18, 24],
                         labels=['Manhã (0-12h)', 'Tarde (12-18h)', 'Noite (18-24h)'])
    
    # 3. Distribuição por Período
    criar_grafico_pizza(
        df['periodo'].value_counts(),
        'Distribuição de Ligações por Período do Dia',
        os.path.join(pasta_graficos, f"periodos_pizza_{data_atual}.png")
    )
    
    # Criar relatório em texto
    with open(os.path.join(pasta_resultados, f"relatorio_{data_atual}.txt"), "w", encoding="utf-8") as f:
        f.write("=== RELATÓRIO DE ANÁLISE DE LIGAÇÕES ===\n\n")
        
        f.write("1. DISTRIBUIÇÃO POR PERÍODO DO DIA\n")
        f.write("-" * 40 + "\n")
        periodo_counts = df['periodo'].value_counts()
        for periodo, count in periodo_counts.items():
            f.write(f"{periodo}: {count} ligações\n")
        f.write("\n")
        
        f.write("2. ANÁLISE POR HORÁRIO E CATEGORIA\n")
        f.write("-" * 40 + "\n")
        for periodo in df['periodo'].unique():
            f.write(f"\n{periodo}:\n")
            df_periodo = df[df['periodo'] == periodo]
            categorias_periodo = df_periodo['categoria'].value_counts()
            for cat, count in categorias_periodo.items():
                f.write(f"  - {cat}: {count} ligações ({(count/len(df_periodo)*100):.1f}%)\n")
        
        f.write("\n3. ANÁLISE DE SENTIMENTO POR PERÍODO\n")
        f.write("-" * 40 + "\n")
        for periodo in df['periodo'].unique():
            f.write(f"\n{periodo}:\n")
            df_periodo = df[df['periodo'] == periodo]
            sentimentos_periodo = df_periodo['sentimento'].value_counts()
            for sent, count in sentimentos_periodo.items():
                f.write(f"  - {sent}: {count} ligações ({(count/len(df_periodo)*100):.1f}%)\n")

        # Adicionar detalhes de cada ligação ordenada por horário
        f.write("\n4. DETALHES DAS LIGAÇÕES POR HORÁRIO\n")
        f.write("-" * 40 + "\n")
        df_sorted = df.sort_values('horario')
        for _, row in df_sorted.iterrows():
            f.write(f"\nHorário: {row['horario']}\n")
            f.write(f"Arquivo: {row['arquivo']}\n")
            f.write(f"Categoria: {row['categoria']}\n")
            f.write(f"Sentimento: {row['sentimento']}\n")
            f.write("-" * 20 + "\n")

if __name__ == "__main__":
    gerar_relatorio() 