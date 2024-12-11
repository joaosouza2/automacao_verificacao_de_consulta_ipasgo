
# Automação de Verificação de Consultas IPASGO

## Como executar o código?  

Para executar o código, basta abrir o aplicativo no notebook VsCode, terá que abrir também a pasta do arquivo do projeto que está localizado no caminho `C:\Users\SUPERVISÃO ADM\Desktop\RPA_verificação_ipasgo`. 
Para abrir irá: 

   1- Clicar em "Abrir a Pasta..." ou  Ctrl+K após isso +O ficando Ctrl+K+O

      1.1 O caminho do arquivo: `C:\Users\SUPERVISÃO ADM\Desktop\RPA_verificação_ipasgo`   


   2- Encontrar o arquivo `main.py`, clicar no play no canto superior direito. 

      2.2 aperte Ctrl+j para abrir o terminal. (todas as respostas de execução do arquivo será feito pelo terminal)

      2.3 existem 2 tipos de terminal, sendo nomeados por BASH e PYTHON. Toda vez que abrir o terminal pelo comando Ctrl+j, será automaticamente no terminal BASH. O terminal PYTHON apenas será trocado quando executar o código. O código ele funciona usando o terminal PYTHON, logo quando clicar no play, irá automaticamente abrir nesse terminal. Mas para executar o terminal precisa estar selecionado no BASH ao invés do PYTHON, é confuso, mas é como o código funciona. Logo toda vez que executar o código e quiser executar novamente, terá que fechar o terminal PYTHON, passe a seta do mouse em cima do nome terminal PYTHON que irá aparecer um `X` para que possa fechar o terminal e mudará autoamticamente para o terminal BASH. Outra solução é apenas clicar no "+" no canto inferior direito do terminal, que abrirá um BASH ou PYTHON caso queira selecionar na seta ao lado do "+"

      





## Estrutura do Projeto

- `main.py`: É o código principal para que possa rodar a automação, esse é o código que deve-rá ser executado.
- `Base_confirmação.xlsx`: É o arquivo que será executado todos os dias com o endereço `C:\Users\SUPERVISÃO ADM\Desktop\RPA_verificação_ipasgo\planilhas\Base_confirmação.xlsx`

## Tecnologias Utilizadas

- **Python 3.x**
- **Selenium WebDriver**
- **Pandas**

## Funcionalidades

- Login automático no sistema IPASGO.
- Preenchimento automático de formulários.
- Verificação de consultas realizadas.
- Geração de logs para monitoramento de erros e execução.

## Como Executar

1. **Clone o Repositório:**
   ```bash
   git clone https://github.com/joaosouza2/automacao_verificacao_de_consulta_ipasgo.git
   ```
2. **Crie e Ative o Ambiente Virtual:**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Linux/Mac
   .venv\Scripts\activate  # Windows
   ```
3. **Instale as Dependências:**
   ```bash
   pip install -r requirements.txt
   ```
4. **Execute o Script:**
   ```bash
   python version_two_verificacao_ipasgo.py
   ```

## Configurações Necessárias

- **WebDriver:** Certifique-se de configurar corretamente o caminho do ChromeDriver no script colocando o arquivo webdriver do selenium no variável de ambiente do windows.
- **Credenciais:** Atualize o arquivo de configurações com suas credenciais do sistema IPASGO.

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais informações.

## Contato

- **Autor:** João Pedro Souza
- **Email:** joaosouza2@disce.ufg.br
- **GitHub:** [joaosouza2](https://github.com/joaosouza2)
