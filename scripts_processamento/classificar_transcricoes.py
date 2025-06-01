import pandas as pd
import os
import json
from typing import List, Dict, Tuple
import re
from datetime import datetime
from tqdm import tqdm

# Definição das categorias e palavras-chave
CATEGORIAS = {
    'Agendamento de consulta': ['agendar', 'marcar', 'consulta', 'horário', 'disponibilidade'],
    'Reagendamento ou cancelamento': ['remarcar', 'cancelar', 'desmarcar', 'adiar', 'mudança'],
    'Dúvidas sobre cirurgia': ['cirurgia', 'risco', 'preparo', 'procedimento', 'pré-operatório'],
    'Encaminhamento para WhatsApp': ['whatsapp', 'link', 'mensagem', 'zap', 'número'],
    'Solicitação de atestado': ['atestado', 'documento', 'declaração', 'comprovante'],
    'Retorno pós-operatório': ['retorno', 'pós', 'avaliação', 'pós-operatório', 'acompanhamento'],
    'Cobranças e valores': ['valor', 'preço', 'pagamento', 'custo', 'orçamento', 'parcela'],
    'Ligação indevida/sem resposta': ['alô', 'barulho', 'nada', 'silêncio', 'ruído'],
    'Reclamação': ['reclamação', 'insatisfação', 'problema', 'queixa', 'insatisfeito', 'reclamar']
}

def carregar_transcricoes(pasta_transcricoes: str) -> List[Dict]:
    """
    Carrega todas as transcrições da pasta especificada.
    Retorna uma lista de dicionários com os dados de cada transcrição
    """
    transcricoes = []
    
    # Verifica se a pasta existe
    if not os.path.exists(pasta_transcricoes):
        print(f"Pasta {pasta_transcricoes} não encontrada!")
        return transcricoes
    
    # Lista todos os arquivos de transcrição
    arquivos = [f for f in os.listdir(pasta_transcricoes) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    print(f"Encontrados {len(arquivos)} arquivos de transcrição.")
    
    # Carrega cada arquivo
    for arquivo in tqdm(arquivos, desc="Carregando transcrições"):
        caminho_completo = os.path.join(pasta_transcricoes, arquivo)
        try:
            # Carrega o arquivo Excel individual
            df = pd.read_excel(caminho_completo)
            if 'nome_arquivo' in df.columns and 'texto_transcrito' in df.columns:
                dados = df.iloc[0].to_dict()  # Converte a primeira linha para dicionário
                transcricoes.append(dados)
        except Exception as e:
            print(f"Erro ao carregar {arquivo}: {str(e)}")
    
    return transcricoes

def detectar_categorias(texto: str) -> Tuple[List[str], Dict[str, str]]:
    """
    Detecta as categorias presentes no texto e retorna os trechos relevantes.
    Retorna: (lista_categorias, dict_trechos_representativos)
    """
    categorias_detectadas = []
    trechos = {}
    
    if not isinstance(texto, str):
        return ["Erro na transcrição"], {"Erro na transcrição": "Texto não disponível"}
    
    # Converte o texto para minúsculas para comparação
    texto_lower = texto.lower()
    
    # Para cada categoria e suas palavras-chave
    for categoria, palavras in CATEGORIAS.items():
        for palavra in palavras:
            if palavra.lower() in texto_lower:
                # Encontra o trecho do texto que contém a palavra-chave
                padrao = re.compile(f'.{{0,50}}{palavra}.{{0,50}}', re.IGNORECASE)
                match = padrao.search(texto)
                if match:
                    trecho = match.group().strip()
                    # Adiciona a categoria se ainda não foi detectada
                    if categoria not in categorias_detectadas:
                        categorias_detectadas.append(categoria)
                        trechos[categoria] = trecho
    
    # Se nenhuma categoria foi detectada, classifica como "Outros assuntos"
    if not categorias_detectadas:
        categorias_detectadas.append("Outros assuntos")
        trechos["Outros assuntos"] = texto[:100] + "..."  # Primeiros 100 caracteres
    
    return categorias_detectadas, trechos

def classificar_transcricoes():
    """
    Função principal que coordena o processo de classificação.
    """
    # Configuração de pastas
    pasta_transcricoes = os.path.join("scripts_processamento", "bd_transcricao", "2025-05", "transcricoes_individuais")
    pasta_resultados = os.path.join("scripts_processamento", "bd_transcricao", "2025-05")
    
    print("Iniciando classificação das transcrições disponíveis...")
    
    # Carrega as transcrições
    transcricoes = carregar_transcricoes(pasta_transcricoes)
    
    if not transcricoes:
        print("Nenhuma transcrição encontrada para processar!")
        return
    
    # Cria DataFrame com as transcrições carregadas
    df = pd.DataFrame(transcricoes)
    
    print(f"\nProcessando {len(df)} transcrições...")
    
    # Adiciona as novas colunas para categorização
    df['categorias_detectadas'] = None
    df['trecho_representativo'] = None
    
    # Processa cada transcrição
    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Classificando"):
        texto = row['texto_transcrito']
        categorias, trechos = detectar_categorias(texto)
        
        # Atualiza o DataFrame
        df.at[idx, 'categorias_detectadas'] = ', '.join(categorias)
        df.at[idx, 'trecho_representativo'] = json.dumps(trechos, ensure_ascii=False)
    
    # Salva o resultado
    arquivo_saida = os.path.join(pasta_resultados, f'classificacao_parcial_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    df.to_excel(arquivo_saida, index=False)
    
    print("\nClassificação parcial concluída!")
    print(f"Resultados salvos em: {arquivo_saida}")
    
    # Exibe estatísticas das categorias
    print("\nEstatísticas por categoria:")
    for categoria in list(CATEGORIAS.keys()) + ["Outros assuntos"]:
        total = df['categorias_detectadas'].str.contains(categoria, na=False).sum()
        percentual = (total / len(df)) * 100
        print(f"{categoria}: {total} ligações ({percentual:.1f}%)")
    
    print("\nObservação: Esta é uma classificação parcial das transcrições já disponíveis.")
    print("Você pode executar novamente este script mais tarde para classificar mais transcrições.")

if __name__ == "__main__":
    classificar_transcricoes() 