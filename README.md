
# Automação de Verificação de Consultas IPASGO

## Descrição

Este projeto automatiza a verificação de consultas no sistema IPASGO utilizando Selenium WebDriver. O código preenche formulários, faz login no sistema e interage com elementos da interface para registrar e verificar consultas automaticamente.

## Estrutura do Projeto

- `main.py`: Código principal que executa a automação.
- `Base_confirmação.xlsx`: Arquivo Excel usado para preencher campos automaticamente. (a pasta está dentro da pasta PLANILHAS `C:\Users\SUPERVISÃO ADM\Desktop\RPA_verificação_ipasgo\planilhas`)

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
- **Credenciais:** Atualize o arquivo de configurações com suas credenciais do sistema IPASGO. Pode-se ser definido pela variável de ambiente permanente ou temporária.

## Licença

Este projeto está sob a licença MIT. Consulte o arquivo LICENSE para mais informações.

## Contato

- **Autor:** João Pedro Souza
- **Email:** joaosouza2@disce.ufg.br
- **GitHub:** [joaosouza2](https://github.com/joaosouza2)
