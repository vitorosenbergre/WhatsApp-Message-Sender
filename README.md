
# WhatsApp Message Sender

Este script automatiza o envio de mensagens pelo WhatsApp Web usando Python. Ele utiliza uma planilha Excel para obter os dados dos clientes e gera mensagens personalizadas com informações como nome e data de vencimento.

## Funcionalidades

- Carrega dados de clientes de um arquivo Excel.
- Gera mensagens personalizadas com base nas informações da planilha.
- Envia as mensagens automaticamente pelo WhatsApp Web.
- Registra erros em um arquivo CSV caso não seja possível enviar mensagens para determinados clientes.

## Requisitos

- **Python 3.8 ou superior**
- **Bibliotecas Python:**
  - openpyxl
  - webbrowser
  - pyautogui

## Como Usar

1. **Preparação dos Dados:**

   - Crie um arquivo Excel (`assets/clientes.xlsx`) com uma planilha chamada `Planilha1`.
   - Certifique-se de que a planilha tem as seguintes colunas:
     - **Nome** (coluna A): Nome do cliente.
     - **Telefone** (coluna B): Telefone do cliente no formato internacional (e.g., `5511999999999`).
     - **Vencimento** (coluna C): Data de vencimento do boleto (formato `AAAA-MM-DD`).

2. **Configuração Inicial:**

   - Certifique-se de que o navegador padrão está configurado para abrir links e que o WhatsApp Web está ativo.

3. **Executar o Script:**

   - Rode o script com o comando:
     ```bash
     python script.py
     ```

4. **Verificação de Erros:**

   - Caso algum cliente não receba a mensagem, os detalhes serão registrados em `assets/erros.csv`.

## Estrutura de Arquivos

```
assets/
  clientes.xlsx   # Planilha com os dados dos clientes
  erros.csv       # Arquivo gerado para registrar erros
script.py         # Script principal
```

## Observações

- Certifique-se de que as imagens utilizadas no código (e.g., `enter.png`) estejam localizadas no diretório `assets`.
- O envio das mensagens depende da estabilidade do WhatsApp Web e da conexão à internet.
- Ajuste os tempos de espera (`sleep`) caso o desempenho do seu computador ou internet seja diferente do esperado.

## Limitações

- Este script não garante 100% de entrega devido a erros de rede ou bloqueios no WhatsApp Web.
- O uso excessivo pode levar ao bloqueio temporário da conta pelo WhatsApp.

## Autor

Desenvolvido por [Seu Nome].
