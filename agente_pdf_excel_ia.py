import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import json
import os
import pypdf
import pandas as pd
import google.generativeai as genai

# Define o nome do arquivo que guardará a chave da API
CONFIG_FILE = "config.json"

class AgentPDFtoExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Agente IA - Conversor de Relatório (PDF -> Excel)")
        self.root.geometry("520x280")
        self.root.resizable(False, False)
        
        # Configuração de Estilo Simples
        style = ttk.Style()
        style.theme_use('clam')
        
        # Variáveis de Controle
        self.pdf_path = tk.StringVar()
        self.api_key = tk.StringVar()
        
        # Carrega a chave da API salva (se existir) antes de montar a interface
        self.carregar_configuracao()
        
        self.create_widgets()

    def carregar_configuracao(self):
        """Lê o arquivo config.json e preenche a variável da API Key."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    config = json.load(f)
                    self.api_key.set(config.get("api_key", ""))
            except Exception:
                pass # Se houver erro de leitura, apenas ignora e deixa o campo vazio

    def salvar_configuracao(self):
        """Salva a API Key digitada no arquivo config.json."""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"api_key": self.api_key.get().strip()}, f)
        except Exception as e:
            print(f"Erro ao salvar configuração: {e}")

    def create_widgets(self):
        # Frame Principal
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Campo da Chave de API
        ttk.Label(main_frame, text="Google Gemini API Key:").grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        api_entry = ttk.Entry(main_frame, textvariable=self.api_key, width=54, show="*")
        api_entry.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(0, 15))

        # Seleção de Arquivo PDF
        ttk.Label(main_frame, text="Arquivo PDF (Relatório):").grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(0, 20))
        
        pdf_entry = ttk.Entry(file_frame, textvariable=self.pdf_path, width=42, state='readonly')
        pdf_entry.pack(side=tk.LEFT, padx=(0, 10))
        
        btn_browse = ttk.Button(file_frame, text="Procurar...", command=self.browse_file)
        btn_browse.pack(side=tk.LEFT)

        # Botão de Ação e Barra de Progresso
        self.btn_convert = ttk.Button(main_frame, text="Converter para Excel", command=self.start_conversion_thread)
        self.btn_convert.grid(row=4, column=0, columnspan=2, pady=(10, 10))

        self.status_label = ttk.Label(main_frame, text="", foreground="gray")
        self.status_label.grid(row=5, column=0, columnspan=2)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Selecione o Relatório em PDF",
            filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if filename:
            self.pdf_path.set(filename)

    def start_conversion_thread(self):
        if not self.api_key.get().strip():
            messagebox.showwarning("Aviso", "Por favor, insira a chave da API do Gemini.")
            return
        if not self.pdf_path.get().strip():
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF.")
            return

        # Salva a chave da API informada para as próximas execuções
        self.salvar_configuracao()

        # Desabilita o botão para evitar múltiplos cliques
        self.btn_convert.config(state=tk.DISABLED)
        self.status_label.config(text="Lendo PDF e processando com IA... Aguarde.", foreground="blue")
        
        # Inicia a thread para não congelar a GUI
        thread = threading.Thread(target=self.process_conversion)
        thread.start()

    def process_conversion(self):
        try:
            # 1. Extração do texto do PDF
            texto_pdf = self.extract_text_from_pdf(self.pdf_path.get())
            if not texto_pdf.strip():
                raise ValueError("Não foi possível extrair texto legível deste PDF.")

            # 2. Chamada ao Agente de IA para estruturar os dados
            dados_estruturados = self.call_ai_agent(texto_pdf, self.api_key.get().strip())

            # 3. Conversão para Excel (xlsx)
            output_path = self.pdf_path.get().replace(".pdf", "_convertido.xlsx")
            self.save_to_excel(dados_estruturados, output_path)

            # Sucesso
            self.root.after(0, self.conversion_success, output_path)

        except Exception as e:
            # Erro
            self.root.after(0, self.conversion_error, str(e))

    def extract_text_from_pdf(self, path):
        text = ""
        with open(path, "rb") as file:
            reader = pypdf.PdfReader(file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text

    def call_ai_agent(self, text, api_key):
        genai.configure(api_key=api_key)
        
        # O modelo flash tem janela de contexto ampla e é eficiente
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Você é um agente de software estritamente lógico e especializado em extração de dados.
        O texto abaixo foi extraído de um relatório em PDF. Sua única função é analisar este texto, 
        identificar os dados tabulares ou estruturados inerentes ao relatório, e retornar o resultado 
        como um array JSON válido.

        Regras absolutas:
        1. O retorno DEVE ser APENAS um array JSON `[]`. Nenhuma palavra adicional, sem introdução.
        2. Cada objeto dentro do array representará uma linha da planilha do Excel.
        3. Normalize os nomes das chaves (nomes das colunas) para serem claros e descritivos.
        4. Não inclua blocos de formatação markdown (como ```json). Apenas o JSON puro.

        Texto do Relatório:
        {text}
        """
        
        response = model.generate_content(prompt)
        raw_output = response.text.strip()

        # Limpeza de segurança caso a IA retorne formatação markdown por engano
        if raw_output.startswith("```json"):
            raw_output = raw_output[7:-3]
        elif raw_output.startswith("```"):
            raw_output = raw_output[3:-3]

        try:
            json_data = json.loads(raw_output.strip())
            return json_data
        except json.JSONDecodeError:
            raise ValueError("O Agente IA não retornou um JSON válido. Verifique se o relatório contém dados estruturáveis.")

    def save_to_excel(self, json_data, output_path):
        # Valida se os dados são uma lista
        if not isinstance(json_data, list):
            # Se for um dicionário único, transforma em lista
            json_data = [json_data]
            
        df = pd.DataFrame(json_data)
        df.to_excel(output_path, index=False, engine='openpyxl')

    def conversion_success(self, output_path):
        self.btn_convert.config(state=tk.NORMAL)
        self.status_label.config(text="Conversão concluída com sucesso!", foreground="green")
        messagebox.showinfo("Sucesso", f"Planilha gerada em:\n{output_path}")

    def conversion_error(self, error_message):
        self.btn_convert.config(state=tk.NORMAL)
        self.status_label.config(text="Erro durante a conversão.", foreground="red")
        messagebox.showerror("Erro de Processamento", f"Ocorreu um erro:\n{error_message}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AgentPDFtoExcelApp(root)
    root.mainloop()