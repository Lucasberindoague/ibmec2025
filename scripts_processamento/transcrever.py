import os
import whisper
import pandas as pd
from tqdm import tqdm
import re
from datetime import datetime
import csv

def extrair_ramal(nome_arquivo):
    """Extrai o ramal do nome do arquivo."""
    padrao = r"bioc(\d{4})"
    match = re.search(padrao, nome_arquivo)
    if match:
        return f"bioc{match.group(1)}"
    return "Não identificado"

def carregar_checkpoint(arquivo_checkpoint):
    """Carrega a lista de arquivos já processados do checkpoint."""
    arquivos_processados = set()
    if os.path.exists(arquivo_checkpoint):
        with open(arquivo_checkpoint, 'r', newline='') as f:
            reader = csv.reader(f)
            next(reader)  # Pula o cabeçalho
            for row in reader:
                arquivos_processados.add(row[0])
    return arquivos_processados

def salvar_checkpoint(arquivo_checkpoint, nome_arquivo, resultado):
    """Salva o progresso no arquivo de checkpoint."""
    arquivo_existe = os.path.exists(arquivo_checkpoint)
    with open(arquivo_checkpoint, 'a', newline='') as f:
        writer = csv.writer(f)
        if not arquivo_existe:
            writer.writerow(['nome_arquivo', 'timestamp', 'duracao_segundos', 'status'])
        writer.writerow([
            nome_arquivo,
            datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            resultado['duracao'],
            'sucesso' if not resultado['texto'].startswith('ERRO:') else 'erro'
        ])

def salvar_transcricao_individual(pasta_resultados_individuais, nome_arquivo, dados):
    """Salva a transcrição individual em um arquivo Excel."""
    nome_arquivo_excel = f"{os.path.splitext(nome_arquivo)[0]}.xlsx"
    caminho_arquivo = os.path.join(pasta_resultados_individuais, nome_arquivo_excel)
    
    df = pd.DataFrame([dados])
    df.to_excel(caminho_arquivo, index=False)

def transcrever_audio(audio_path):
    """Transcreve um arquivo de áudio usando o Whisper com modelo small."""
    try:
        # Carrega o modelo small (mais rápido)
        model = whisper.load_model("small")
        
        # Realiza a transcrição
        result = model.transcribe(audio_path, language="pt", fp16=False)
        
        return {
            'texto': result["text"],
            'duracao': result["segments"][-1]["end"] if result["segments"] else 0
        }
    except Exception as e:
        print(f"Erro ao transcrever {os.path.basename(audio_path)}: {str(e)}")
        return {
            'texto': f"ERRO: {str(e)}",
            'duracao': 0
        }

def processar_audios():
    """Processa todos os arquivos de áudio na pasta de gravações."""
    # Obtém o diretório raiz do projeto
    diretorio_raiz = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Configuração de pastas
    pasta_audios = os.path.join(diretorio_raiz, "automatizacoes_de_download", "automatizacao_downloads_audios", "ligacoes_descompactadas", "2025-05")
    pasta_resultados = os.path.join(diretorio_raiz, "scripts_processamento", "bd_transcricao", "2025-05")
    pasta_resultados_individuais = os.path.join(pasta_resultados, "transcricoes_individuais")
    arquivo_checkpoint = os.path.join(pasta_resultados, "checkpoint.csv")
    
    # Cria as pastas necessárias
    os.makedirs(pasta_resultados, exist_ok=True)
    os.makedirs(pasta_resultados_individuais, exist_ok=True)
    
    print("Iniciando processamento de áudios com modelo SMALL do Whisper")
    print(f"Os resultados serão salvos na pasta {pasta_resultados}")
    
    # Carrega o checkpoint
    arquivos_processados = carregar_checkpoint(arquivo_checkpoint)
    print(f"Encontrados {len(arquivos_processados)} arquivos já processados")
    
    # Lista todos os arquivos MP3
    arquivos_mp3 = [f for f in os.listdir(pasta_audios) if f.endswith('.mp3')]
    
    if not arquivos_mp3:
        print(f"Nenhum arquivo MP3 encontrado em {pasta_audios}")
        return
    
    # Filtra apenas arquivos não processados
    arquivos_pendentes = [f for f in arquivos_mp3 if f not in arquivos_processados]
    print(f"Arquivos pendentes para processamento: {len(arquivos_pendentes)}")
    
    # Prepara a lista para armazenar os resultados
    resultados = []
    
    # Processa cada arquivo pendente
    for arquivo in tqdm(arquivos_pendentes, desc="Processando áudios"):
        caminho_completo = os.path.join(pasta_audios, arquivo)
        
        # Extrai o ramal
        ramal = extrair_ramal(arquivo)
        
        # Transcreve o áudio
        resultado = transcrever_audio(caminho_completo)
        
        # Prepara os dados da transcrição
        dados_transcricao = {
            'nome_arquivo': arquivo,
            'texto_transcrito': resultado['texto'],
            'duracao_segundos': resultado['duracao'],
            'caminho_arquivo': caminho_completo,
            'ramal': ramal,
            'data_processamento': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # Salva a transcrição individual
        salvar_transcricao_individual(pasta_resultados_individuais, arquivo, dados_transcricao)
        
        # Atualiza o checkpoint
        salvar_checkpoint(arquivo_checkpoint, arquivo, resultado)
        
        # Adiciona à lista de resultados
        resultados.append(dados_transcricao)
    
    # Se houver resultados novos, atualiza o arquivo consolidado
    if resultados:
        # Carrega dados existentes se houver
        arquivo_consolidado = os.path.join(pasta_resultados, "transcricoes_consolidadas.xlsx")
        if os.path.exists(arquivo_consolidado):
            df_existente = pd.read_excel(arquivo_consolidado)
            df_novos = pd.DataFrame(resultados)
            df_final = pd.concat([df_existente, df_novos], ignore_index=True)
        else:
            df_final = pd.DataFrame(resultados)
        
        # Salva o arquivo consolidado atualizado
        df_final.to_excel(arquivo_consolidado, index=False)
    
    print(f"\nProcessamento concluído!")
    print(f"Total de arquivos processados nesta execução: {len(resultados)}")
    print(f"Total de arquivos processados (incluindo anteriores): {len(arquivos_processados) + len(resultados)}")
    print(f"Resultados individuais salvos em: {pasta_resultados_individuais}")
    print(f"Arquivo consolidado salvo em: {pasta_resultados}")

if __name__ == "__main__":
    processar_audios() 