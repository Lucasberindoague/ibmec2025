import os
import zipfile
from pathlib import Path
import datetime

class DescompactadorGravacoes:
    def __init__(self):
        # Configura os diretórios
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.downloads_dir = os.path.join(self.base_dir, "downloads")  # Pasta correta dos ZIPs
        self.destino_dir = os.path.join(self.base_dir, "ligacoes_descompactadas")
        
        # Contadores para o resumo
        self.total_arquivos_audio = 0
        self.subdiretorios_criados = set()
    
    def encontrar_arquivos_zip(self):
        """Encontra todos os arquivos .zip de gravações na pasta downloads"""
        arquivos_zip = []
        for arquivo in os.listdir(self.downloads_dir):
            if arquivo.lower().endswith('.zip') and "RECORDINGS_BULK" in arquivo:
                caminho_completo = os.path.join(self.downloads_dir, arquivo)
                arquivos_zip.append(caminho_completo)
        return arquivos_zip
    
    def descompactar_arquivo(self, arquivo_zip):
        """Descompacta um arquivo zip mantendo sua estrutura"""
        try:
            with zipfile.ZipFile(arquivo_zip, 'r') as zip_ref:
                # Lista todos os arquivos dentro do ZIP
                for arquivo in zip_ref.namelist():
                    # Extrai o arquivo
                    zip_ref.extract(arquivo, self.destino_dir)
                    
                    # Atualiza contadores
                    if arquivo.lower().endswith(('.wav', '.mp3', '.ogg')):
                        self.total_arquivos_audio += 1
                    
                    # Registra subdiretórios
                    diretorio = os.path.dirname(arquivo)
                    if diretorio:
                        self.subdiretorios_criados.add(diretorio)
                
                print(f"Arquivo descompactado: {os.path.basename(arquivo_zip)}")
                
        except Exception as e:
            print(f"Erro ao descompactar {os.path.basename(arquivo_zip)}: {str(e)}")
    
    def processar_arquivos(self):
        """Processa todos os arquivos ZIP encontrados"""
        print("\n=== Descompactador de Gravações ===")
        print(f"Procurando arquivos ZIP em: {self.downloads_dir}")
        print(f"Destino dos arquivos: {self.destino_dir}\n")
        
        # Garante que a pasta de destino existe
        os.makedirs(self.destino_dir, exist_ok=True)
        
        # Encontra e processa os arquivos
        arquivos_zip = self.encontrar_arquivos_zip()
        
        if not arquivos_zip:
            print("Nenhum arquivo ZIP de gravação encontrado!")
            return
        
        print(f"Encontrados {len(arquivos_zip)} arquivos ZIP de gravações\n")
        
        # Descompacta cada arquivo
        for arquivo in arquivos_zip:
            self.descompactar_arquivo(arquivo)
        
        # Imprime o resumo
        print("\n=== Resumo da Extração ===")
        print(f"Total de arquivos de áudio extraídos: {self.total_arquivos_audio}")
        
        if self.subdiretorios_criados:
            print("\nSubdiretórios criados:")
            for subdir in sorted(self.subdiretorios_criados):
                print(f"- {subdir}")
        else:
            print("\nTodos os arquivos foram extraídos na raiz da pasta de destino.")

if __name__ == "__main__":
    descompactador = DescompactadorGravacoes()
    descompactador.processar_arquivos() 