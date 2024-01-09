# Automacao-Excel-Python

Bem-vindo ao projeto Automacao-Excel-Python! Este é um script de automação desenvolvido em Python utilizando a biblioteca PyAutoGUI, 
destinado a interagir com uma planilha do Excel. O código realiza uma série de ações automatizadas como se estivesse em um sistema,
simulando o comportamento do usuário ao preencher campos de login e senha, clicar em botões e executar operações como voltar e apagar dados.

## Funcionalidades Principais

- Abre a ferramenta do Excel a partir do menu Iniciar do Windows.
- Preenche os campos de login e senha com informações pré-definidas.
- Simula o clique nos botões de login e voltar.
- Apaga os dados inseridos nos campos de login e senha.

## Instruções de Uso

1. Execute o script em um ambiente Python.
2. Aguarde a automação interagir com a planilha do Excel.

## Requisitos

Certifique-se de ter instalado os seguintes pacotes:

```bash
pip install pyautogui


### Comandos PyAutoGUI Utilizados:

- **`pyautogui.press("win")`**: Simula o pressionamento da tecla "Win" (Windows) para abrir o menu Iniciar do Windows.

- **`pyautogui.write("Login.xlsx")`**: Digita o nome do arquivo ("Login.xlsx") para abrir a ferramenta do Excel.

- **`pyautogui.press("enter")`**: Simula a tecla "Enter" para confirmar a abertura do arquivo.

- **`pyautogui.click(x=301, y=301)`**: Clica na posição especificada (301, 301) na tela, preenchendo o campo de login.

- **`pyautogui.write("Lais")`**: Digita o nome de usuário "Lais" no campo de login.

- **`pyautogui.click(x=431, y=340)`**: Clica na posição (431, 340) na tela, movendo o foco para o campo de senha.

- **`pyautogui.write("1234")`**: Digita a senha "1234" no campo de senha.

- **`pyautogui.click(x=409, y=429)`**: Clica na posição (409, 429) na tela, simulando o clique no botão "Fazer Login".

- **`pyautogui.click(x=699, y=572)`**: Clica na posição (699, 572) na tela, simulando o clique no botão "Voltar".

- **`pyautogui.click(x=471, y=294)`**: Clica na posição (471, 294) na tela, movendo o cursor para o campo de login.

- **`pyautogui.press("backspace")`**: Simula o pressionamento da tecla "Backspace" para apagar o conteúdo do campo.

- **`pyautogui.click(x=431, y=340)`**: Clica na posição (431, 340) na tela, movendo o cursor para o campo de senha.

- **`pyautogui.press("backspace")`**: Simula o pressionamento da tecla "Backspace" para apagar o conteúdo do campo de senha.

