import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import os
import json

class RenomeadorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Renomeador de Arquivos Corporativo")
        self.geometry("500x400")
        self.resizable(False, False)
        
        # Estilo para o texto de exemplo (placeholder)
        self.style = ttk.Style(self)
        self.style.configure("Placeholder.TEntry", foreground="gray")

        self.config_file = "renomeador_config.json"

        # Variáveis
        self.setor_selecionado = tk.StringVar()
        self.nome_funcionario = tk.StringVar()
        self.nome_arquivo_original = tk.StringVar()
        self.versao_arquivo = tk.StringVar()
        self.caminho_arquivo_original = tk.StringVar()

        # Dicionário de Setores
        self.setores = {
            "Atendimento": "Ate", "Conteúdo": "Con", "Departamento Pessoal": "Dep",
            "Eventos": "Eve", "Financeiro": "Fin", "Gerência": "Ger", "Mentorias": "Men"
        }

        # <<< 1. Lista de funcionários em ordem alfabética
        self.funcionarios = sorted([
            "Flávio", "Thiago", "Moises", "Mateus", "Sara", "João", "Luiza",
            "Patrick", "Paula", "Paulo", "Yarhima", "Mayara", "Clara",
            "Renilson", "Lane", "Juliana", "Deisy", "JP", "Glenda",
            "Rayssa", "Walkyria"
        ])

        # Componentes da UI
        self.criar_widgets()
        self.carregar_configuracoes()

    def carregar_configuracoes(self):
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
                setor = config_data.get('ultimo_setor', "Selecione o setor")
                funcionario = config_data.get('ultimo_funcionario', "Selecione o funcionário")
                self.setor_selecionado.set(setor)
                self.nome_funcionario.set(funcionario)
        except (FileNotFoundError, json.JSONDecodeError):
            self.setor_selecionado.set("Selecione o setor")
            self.nome_funcionario.set("Selecione o funcionário")

    def salvar_configuracoes(self):
        config_data = {
            'ultimo_setor': self.setor_selecionado.get(),
            'ultimo_funcionario': self.nome_funcionario.get()
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config_data, f, indent=4)

    def _setup_placeholder(self, entry, text):
        entry.insert(0, text)
        entry.configure(style="Placeholder.TEntry")
        
        def on_focus_in(event):
            if entry.get() == text:
                entry.delete(0, tk.END)
                entry.configure(style="TEntry")

        def on_focus_out(event):
            if not entry.get():
                entry.insert(0, text)
                entry.configure(style="Placeholder.TEntry")

        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def criar_widgets(self):
        frame = ttk.Frame(self, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="Data:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.data_label = ttk.Label(frame, text=self.obter_data_atual())
        self.data_label.grid(row=0, column=1, sticky=tk.W, pady=5)

        ttk.Label(frame, text="Setor:").grid(row=1, column=0, sticky=tk.W, pady=5)
        setor_dropdown = ttk.Combobox(frame, textvariable=self.setor_selecionado, values=list(self.setores.keys()), state="readonly")
        setor_dropdown.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)

        # <<< 2. Campo de funcionário agora é um dropdown
        ttk.Label(frame, text="Nome do Funcionário:").grid(row=2, column=0, sticky=tk.W, pady=5)
        nome_funcionario_dropdown = ttk.Combobox(frame, textvariable=self.nome_funcionario, values=self.funcionarios, state="readonly")
        nome_funcionario_dropdown.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)

        # <<< 3. Campos de texto com placeholder
        ttk.Label(frame, text="Nome do Documento:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.nome_arquivo_entry = ttk.Entry(frame, textvariable=self.nome_arquivo_original)
        self.nome_arquivo_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        self._setup_placeholder(self.nome_arquivo_entry, "Ex: ADITIVO CONTRATUAL")

        ttk.Label(frame, text="Versão (só números):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.versao_entry = ttk.Entry(frame, textvariable=self.versao_arquivo)
        self.versao_entry.grid(row=4, column=1, sticky=(tk.W, tk.E), pady=5)
        self._setup_placeholder(self.versao_entry, "Ex: 1")

        ttk.Label(frame, text="Arquivo Original:").grid(row=5, column=0, sticky=tk.W, pady=5)
        selecionar_button = ttk.Button(frame, text="Selecionar Arquivo", command=self.selecionar_arquivo)
        selecionar_button.grid(row=5, column=1, sticky=(tk.W, tk.E), pady=5)
        self.arquivo_selecionado_label = ttk.Label(frame, text="Nenhum arquivo selecionado", foreground="gray")
        self.arquivo_selecionado_label.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=2)

        renomear_button = ttk.Button(frame, text="Renomear Arquivo", command=self.renomear_arquivo)
        renomear_button.grid(row=7, column=0, columnspan=2, pady=20)

        frame.columnconfigure(1, weight=1)

    def obter_data_atual(self):
        # A data é sempre a data atual do sistema
        return datetime.date.today().strftime("%Y%m%d")

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename()
        if caminho:
            self.caminho_arquivo_original.set(caminho)
            nome_arquivo = os.path.basename(caminho)
            self.arquivo_selecionado_label.config(text=nome_arquivo, foreground="black")

    def renomear_arquivo(self):
        # <<< 4. Validações atualizadas para os menus dropdown
        setor = self.setor_selecionado.get()
        if not setor or setor == "Selecione o setor":
            messagebox.showerror("Erro de Validação", "Por favor, selecione um setor.")
            return

        funcionario = self.nome_funcionario.get()
        if not funcionario or funcionario == "Selecione o funcionário":
            messagebox.showerror("Erro de Validação", "Por favor, selecione o nome do funcionário.")
            return
            
        if not self.nome_arquivo_original.get().strip() or self.nome_arquivo_original.get() == "Ex: ADITIVO CONTRATUAL":
            messagebox.showerror("Erro de Validação", "Por favor, insira o nome do documento.")
            return

        versao_str = self.versao_arquivo.get()
        if not versao_str.strip() or not versao_str.isdigit():
            messagebox.showerror("Erro de Validação", "Por favor, insira um número válido para a versão.")
            return

        if not self.caminho_arquivo_original.get():
            messagebox.showerror("Erro de Validação", "Por favor, selecione o arquivo a ser renomeado.")
            return

        try:
            self.salvar_configuracoes()

            data = self.obter_data_atual()
            setor_abrev = self.setores[setor]
            nome_doc = self.nome_arquivo_original.get().strip().upper().replace(" ", "")
            versao = f"v{int(versao_str):02d}"
            
            caminho_original = self.caminho_arquivo_original.get()
            diretorio = os.path.dirname(caminho_original)
            extensao = os.path.splitext(caminho_original)[1]

            novo_nome = f"{data}-{setor_abrev}-{funcionario}-{nome_doc}-{versao}{extensao}"
            novo_caminho = os.path.join(diretorio, novo_nome)

            os.rename(caminho_original, novo_caminho)
            messagebox.showinfo("Sucesso", f"Arquivo renomeado com sucesso para:\n{novo_nome}")

            # Limpa e restaura placeholders
            self.nome_arquivo_entry.delete(0, tk.END)
            self.versao_entry.delete(0, tk.END)
            self._setup_placeholder(self.nome_arquivo_entry, "Ex: ADITIVO CONTRATUAL")
            self._setup_placeholder(self.versao_entry, "Ex: 1")
            
            self.caminho_arquivo_original.set("")
            self.arquivo_selecionado_label.config(text="Nenhum arquivo selecionado", foreground="gray")

        except Exception as e:
            messagebox.showerror("Erro na Renomeação", f"Ocorreu um erro ao renomear o arquivo:\n{e}")

if __name__ == "__main__":
    app = RenomeadorApp()
    app.mainloop()