import pandas as pd
import os
from datetime import datetime

class ClassificadorLigacoes:
    def __init__(self):
        self.mapeamento_atendentes = {
            'bioc5318': 'Joelma',
            'bioc5319': 'Joelma',
            'bioc5316': 'Bia',
            'bioc5310': 'Lucas',
            'bioc5311': 'Ana Rafaela',
            'bioc5313': 'Júlio',
            'bioc5315': 'Lucila'
        }
    
    def extrair_info_arquivo(self, nome_arquivo):
        """Extrai informações do nome do arquivo de áudio"""
        # Remove a extensão
        nome = os.path.splitext(nome_arquivo)[0]
        
        # Divide o nome em partes
        partes = nome.split('_')
        
        # Extrai data e hora
        data = f"{partes[0]}-{partes[1]}-{partes[2]}"
        hora = f"{partes[3]}:{partes[4]}:{partes[5]}"
        
        # Extrai ramal e identifica atendente
        ramal = partes[6]
        atendente = self.mapeamento_atendentes.get(ramal, 'Não identificado')
        
        return {
            'arquivo': nome_arquivo,
            'data': data,
            'hora': hora,
            'ramal': ramal,
            'atendente': atendente,
            'duracao_segundos': None  # Será preenchido depois
        }
    
    def classificar_ligacoes(self, pasta_audios):
        """Classifica todas as ligações da pasta"""
        dados = []
        
        # Lista todos os arquivos de áudio
        for arquivo in os.listdir(pasta_audios):
            if arquivo.endswith(('.mp3', '.wav')):
                info = self.extrair_info_arquivo(arquivo)
                dados.append(info)
        
        # Cria DataFrame
        df = pd.DataFrame(dados)
        
        # Ordena por data e hora
        df['data_hora'] = pd.to_datetime(df['data'] + ' ' + df['hora'])
        df = df.sort_values('data_hora')
        
        return df
    
    def salvar_classificacao(self, df, caminho_saida):
        """Salva a classificação em um arquivo Excel"""
        # Garante que o diretório de saída existe
        os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
        
        df.to_excel(caminho_saida, index=False)
        print(f"\nClassificação salva em: {caminho_saida}")
        
        # Imprime resumo
        print("\n=== Resumo da Classificação ===")
        print(f"Total de ligações: {len(df)}")
        print("\nLigações por atendente:")
        print(df['atendente'].value_counts())
        print("\nLigações por ramal:")
        print(df['ramal'].value_counts())

if __name__ == "__main__":
    classificador = ClassificadorLigacoes()
    
    # Pasta com os áudios
    pasta_audios = os.path.join("automatizacoes_de_download", "automatizacao_downloads_audios", "ligacoes_descompactadas", "2025-05")
    
    # Classifica as ligações
    df = classificador.classificar_ligacoes(pasta_audios)
    
    # Pasta para salvar os resultados
    pasta_resultados = "resultados"
    os.makedirs(pasta_resultados, exist_ok=True)
    
    # Salva o resultado
    classificador.salvar_classificacao(df, os.path.join(pasta_resultados, "classificacao_ligacoes.xlsx")) 