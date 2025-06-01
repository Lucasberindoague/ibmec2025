from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time

class AutomacaoDownloads:
    def __init__(self):
        """Inicializa a automação configurando os diretórios e opções do Chrome"""
        # Configura os diretórios
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.download_dir = os.path.join(self.base_dir, "downloads")
        
        # Garante que a pasta de downloads existe
        os.makedirs(self.download_dir, exist_ok=True)
        
        # Porta de debug do Chrome
        self.debug_port = "9222"

    def conectar_chrome(self):
        """Conecta ao Chrome existente"""
        try:
            options = Options()
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{self.debug_port}")
            
            # Usa o ChromeDriverManager para gerenciar o driver
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            
            print("Conectado ao Chrome com sucesso!")
            return driver
        except Exception as e:
            print(f"Erro ao conectar ao Chrome: {str(e)}")
            print("\nCertifique-se de que:")
            print("1. O Chrome está aberto com a flag --remote-debugging-port=9222")
            print("2. Execute o arquivo 'abrir_chrome.command' se necessário")
            raise

    def baixar_gravacoes(self):
        """Executa o processo de download das gravações"""
        driver = self.conectar_chrome()
        wait = WebDriverWait(driver, 10)  # Espera explícita de até 10 segundos
        
        try:
            # Aguarda os botões de download ficarem visíveis
            print("\nProcurando botões de download na página atual...")
            botoes = wait.until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "btnRecDownload"))
            )
            
            print(f"Encontrados {len(botoes)} botões de download.")
            
            # Itera sobre os botões e faz o download
            for i, botao in enumerate(botoes, 1):
                try:
                    # Rola até o botão ficar visível
                    driver.execute_script("arguments[0].scrollIntoView(true);", botao)
                    time.sleep(0.5)  # Pequena pausa para a rolagem completar
                    
                    # Clica no botão
                    botao.click()
                    print(f"Download {i}/{len(botoes)} iniciado.")
                    
                    # Aguarda 3 segundos entre os downloads
                    time.sleep(3)
                    
                except Exception as e:
                    print(f"Erro ao clicar no botão {i}: {str(e)}")
            
            print("\nTodos os downloads foram iniciados!")
            print(f"Os arquivos estão sendo salvos em: {self.download_dir}")
            
            # Aguarda mais alguns segundos para os últimos downloads iniciarem
            time.sleep(5)
            
        except Exception as e:
            print(f"Erro durante a automação: {str(e)}")
        finally:
            # Não fecha o navegador ao finalizar
            pass

if __name__ == "__main__":
    print("=== Automação de Downloads de Gravações ===")
    print("Este script irá se conectar ao Chrome existente e baixar as gravações.")
    print("\nANTES DE COMEÇAR:")
    print("1. Certifique-se de que o Chrome foi aberto com o arquivo 'abrir_chrome.command'")
    print("2. Navegue até a página de gravações e faça login se necessário")
    print("3. Os downloads serão salvos em: automatizacao_downloads_audios/downloads/")
    input("\nPressione ENTER quando estiver pronto...")
    
    automacao = AutomacaoDownloads()
    automacao.baixar_gravacoes() 