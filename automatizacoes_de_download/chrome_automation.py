from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import os
import time

def iniciar_chrome_debug():
    """Inicia o Chrome com a porta de depuração se ainda não estiver rodando"""
    # Caminho do Chrome no MacOS
    chrome_path = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
    debug_port = "9222"
    
    # Verifica se o Chrome já está rodando com debug port
    try:
        options = Options()
        options.debugger_address = f"127.0.0.1:{debug_port}"
        driver = webdriver.Chrome(options=options)
        print("Conectado ao Chrome existente!")
        return driver
    except:
        # Inicia o Chrome com a porta de debug
        cmd = f'"{chrome_path}" --remote-debugging-port={debug_port}'
        subprocess.Popen(cmd, shell=True)
        print("Iniciando nova instância do Chrome...")
        time.sleep(2)  # Aguarda o Chrome iniciar
        
        # Conecta ao Chrome
        options = Options()
        options.debugger_address = f"127.0.0.1:{debug_port}"
        driver = webdriver.Chrome(options=options)
        return driver

def automatizar_chrome():
    # Conecta ao Chrome existente ou inicia um novo
    driver = iniciar_chrome_debug()
    
    try:
        # Exemplo de automação - você pode modificar conforme necessário
        driver.get("https://www.google.com")
        print("Navegando para o Google...")
        time.sleep(2)
        
        # Aqui você pode adicionar mais comandos de automação
        
    except Exception as e:
        print(f"Erro durante a automação: {e}")
    
    # Não fecha o navegador ao finalizar
    # driver.quit()

if __name__ == "__main__":
    automatizar_chrome() 