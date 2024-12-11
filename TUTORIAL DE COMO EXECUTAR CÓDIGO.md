
# Automação de Verificação de Consultas IPASGO

## Como executar o código?  

Para executar o código, basta abrir o aplicativo no notebook VsCode, terá que abrir também a pasta do arquivo do projeto que está localizado no caminho `C:\Users\SUPERVISÃO ADM\Desktop\RPA_verificação_ipasgo`. 
Para abrir irá: 
   1- Clicar em abrir p

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
