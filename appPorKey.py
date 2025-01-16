import openpyxl  # Biblioteca para manipular arquivos Excel.
import webbrowser  # Biblioteca para abrir links no navegador.
import pyautogui  # Biblioteca para automatizar interações com a interface gráfica.

from urllib.parse import quote  # Função para codificar mensagens para URLs.
from time import sleep  # Função para pausar a execução por um intervalo de tempo.

# Carrega o arquivo Excel que contém os dados dos clientes.
workbook = openpyxl.load_workbook('assets/clientes.xlsx')
pagina_clientes = workbook['Planilha1']  # Seleciona a planilha especificada.

# Itera pelas linhas da planilha, iniciando da segunda linha (ignora o cabeçalho).
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value  # Coluna 1: Nome do cliente.
    telefone = linha[1].value  # Coluna 2: Telefone do cliente.
    vencimento = linha[2].value  # Coluna 3: Data de vencimento do boleto.

    # Cria a mensagem personalizada usando os dados do cliente.
    mensagem = f'Olá {nome}, boleto vence dia {vencimento.strftime('%d/%m/%Y')}'
    
    try:
        # Gera o link do WhatsApp com a mensagem e o telefone do cliente.
        link_mensagem_whats = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whats)  # Abre o link no navegador.
        sleep(10)  # Aguarda o carregamento da página do WhatsApp Web.

        # Pressiona a tecla Enter para enviar a mensagem.
        pyautogui.press('enter') 
        sleep(5)  # Aguarda para garantir que a mensagem seja enviada.

        # Fecha a aba do navegador.
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)  # Aguarda para evitar problemas de sincronização.
    except:
        # Caso ocorra algum erro, informa que não foi possível enviar a mensagem.
        print(f'Não foi possível enviar mensagem para {nome}')

        # Registra o erro em um arquivo CSV.
        with open('assets/erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}\n')
