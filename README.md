# Agente IA - Conversor de Relatórios (PDF para Excel)

Este projeto implementa um agente de software em Python com interface gráfica (GUI) nativa projetado para extrair, de forma inteligente, dados tabulares e estruturados de relatórios em PDF, convertendo-os diretamente para planilhas do Excel (`.xlsx`). 

A solução utiliza o modelo **Gemini 1.5 Flash** do Google para interpretação semântica dos dados, permitindo ler matrizes complexas que extratores de texto tradicionais falham em processar.

## 🚀 Funcionalidades

* **Processamento Inteligente:** Utiliza IA generativa para estruturar dados textuais em formato tabular (JSON para DataFrame).
* **Interface Gráfica Simples (Tkinter):** Acessível a usuários finais sem experiência em terminal.
* **Execução Assíncrona:** Implementação de `threading` para garantir que a interface não congele durante a comunicação com a API.
* **Memória de Configuração:** Armazenamento automático da API Key no arquivo local `config.json` para facilitar execuções futuras.
* **Portabilidade:** Disponível em formato executável autônomo (`.exe`) para ambientes Windows.

## 📦 Estrutura do Projeto

* `agente_pdf_excel_ia.py`: Código-fonte principal contendo a lógica de extração e a interface gráfica.
* `requirements.txt`: Lista de dependências e bibliotecas Python necessárias.
* `config.json`: Arquivo gerado automaticamente na primeira execução para salvar a chave da API.
* `agente_pdf_excel_ia.spec`: Arquivo de especificação gerado pelo PyInstaller para a compilação.
* `dist/agente_pdf_excel_ia.exe`: Aplicativo compilado, independente e pronto para uso.

## 💻 Como Usar a Versão Executável (.exe)

Para usuários que desejam apenas utilizar o conversor sem configurar um ambiente de desenvolvimento:

1. Navegue até a pasta `dist` gerada no projeto.
2. Dê um duplo clique no arquivo **`agente_pdf_excel_ia.exe`**. Não é necessário ter o Python instalado na máquina.
3. Na interface gráfica, insira a sua **Google Gemini API Key** (necessário apenas na primeira execução).
4. Clique em **Procurar...** e selecione o relatório em formato PDF.
5. Clique em **Converter para Excel**.
6. O agente fará a leitura e criará um novo arquivo com o sufixo `_convertido.xlsx` no mesmo diretório do PDF original.

## 🛠️ Como Executar o Código-Fonte

### Instalação
1. Crie um ambiente virtual para isolar as dependências: `python -m venv venv`
2. Ative o ambiente virtual: `venv\Scripts\activate` (Windows)
3. Instale as bibliotecas necessárias: `pip install -r requirements.txt`

### Execução
Com o ambiente ativado, execute o script principal: `python agente_pdf_excel_ia.py`

## 🏗️ Como Compilar um Novo Executável
`pyinstaller --noconsole --onefile --windowed agente_pdf_excel_ia.py`

## 🛡️ Segurança e Privacidade
* A chave de API do Google Gemini é salva localmente (`config.json`) na pasta do executável. O arquivo está bloqueado no `.gitignore` e não sobe para este repositório.
* Nenhuma informação do PDF é retida localmente pelo software além do arquivo `.xlsx` gerado.
