# Bot Assistente de Automação de Planilhas

Este projeto é um bot assistente desenvolvido em Python, que realiza a automação do lançamento de dados de um sistema em planilhas Excel. Ele utiliza as bibliotecas pyautogui para interação com a interface gráfica e openpyxl para manipulação e escrita de dados em arquivos Excel.

## Funcionalidades

- Extração de dados de um sistema e inserção automática em planilhas Excel.
- Manipulação de planilhas: criação, edição e salvamento de arquivos .xlsx.
- Interação automatizada com a interface do usuário para coleta de dados ou execução de ações no sistema.

## Tecnologias Utilizadas

- Python 3.x: Linguagem de programação principal do projeto.
- openpyxl: Biblioteca para manipulação de arquivos Excel.
- pyautogui: Biblioteca para automação de tarefas com interação na interface gráfica.

## Instalação

* Clone este repositório
```
git clone git@github.com:lluismarcello333/python---PySheets.git
cd python---PySheets
```
* Crie e ative um ambiente virtual (opcional, mas recomendado)
```
python3 -m venv venv
source venv/bin/activate  # No Windows use: venv\Scripts\activate
```
* Instale as dependências
```
pip install -r requirements.txt
```

## Como Usar

### Configuração

- Certifique-se de que os arquivos de planilhas .xlsx estejam no diretório correto ou especifique o caminho no código.
- Defina as coordenadas da interface gráfica do sistema, se necessário, para o uso do pyautogui.

### Execução

- Para executar o bot assistente, execute o script principal:
```
python main.py
```

### Personalização

- Você pode ajustar os parâmetros de execução e as funções de automação no código para adaptá-los ao sistema ou planilhas que deseja automatizar.
