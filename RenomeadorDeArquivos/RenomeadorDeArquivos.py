#!/usr/bin/env python3
"""
Renomeador de Arquivos Corporativo — compacto e legível.
"""

import datetime
import json
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional
import unicodedata

from PIL import Image, ImageTk


def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel_path)


class RenomeadorArquivosApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self._configurar_janela()
        self._inicializar_variaveis()
        self._carregar_logo()
        self._configurar_icone()
        self._configurar_estilo()
        self._criar_interface()
        self._carregar_configuracoes()
        self.after(200, self._mostrar_tutorial_se_necessario)

    # ---------- Configuração básica ----------
    def _configurar_janela(self) -> None:
        self.title("Instituto Walkyria Fernandes - Renomeador de Arquivos")
        self.configure(bg="#F0FFFC")
        w, h = 550, 700
        self.geometry(f"{w}x{h}+{(self.winfo_screenwidth()-w)//2}+{(self.winfo_screenheight()-h)//2}")
        self.resizable(False, False)

    def _inicializar_variaveis(self) -> None:
        self.config_file = "renomeador_config.json"

        self.setor_selecionado = tk.StringVar()
        self.funcionario_selecionado = tk.StringVar()
        self.evento_selecionado = tk.StringVar()
        self.nome_documento = tk.StringVar()
        self.versao_arquivo = tk.StringVar()
        self.caminho_arquivo = tk.StringVar()

        self.ultimo_diretorio = os.path.expanduser("~")
        self.arquivos_selecionados: List[str] = []
        self.tutorial_v1_shown = False

        self.setores: Dict[str, str] = {
            "Atendimento": "Ate",
            "Conteúdo": "Con",
            "Departamento Pessoal": "Dep",
            "Eventos": "Eve",
            "Financeiro": "Fin",
            "Gerência": "Ger",
            "Mentorias": "Men",
        }

        self.eventos: Dict[str, str] = {
            "FisioSummit": "FS",
            "Mentoria Black": "MB",
            "Mentoria Black Diamond": "MBD",
            "Pós-Graduação": "PG",
            "RCA360": "RCA",
            "Congresso RCA": "CRCA",
            "PowerFisio": "PF",
            "RCA Exclusive": "RCAE",
            "RCA na Prática": "RNP",
        }
        for n in range(10, 31):
            self.eventos[f"LI {n}"] = f"LI{n}"
        for n in range(1, 11):
            self.eventos[f"LI Pós {n}"] = f"LIP{n}"

        self.funcionarios: List[str] = sorted(
            [
                "Flávio",
                "Thiago",
                "Moises",
                "Mateus",
                "Sara",
                "João",
                "Luiza",
                "Patrick",
                "Paula",
                "Paulo",
                "Yarhima",
                "Mayara",
                "Clara",
                "Renilson",
                "Lane",
                "Juliana",
                "Deisy",
                "JP",
                "Glenda",
                "Rayssa",
                "Walkyria",
            ]
        )

    def _carregar_logo(self) -> None:
        self.logo_image = None
        try:
            p = resource_path(os.path.join("assets", "logo.png"))
            img = Image.open(p).resize((520, 110), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)
        except Exception:
            self.logo_image = None

    # ---------- Ícone (taskbar + título) ----------
    def _configurar_icone(self) -> None:
        ico_path = resource_path(os.path.join("assets", "logo.ico"))
        png_path = resource_path(os.path.join("assets", "logo.png"))

        if sys.platform.startswith("win") and os.path.exists(ico_path):
            try:
                self.iconbitmap(default=ico_path)
            except Exception:
                pass
            try:
                import ctypes
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("WF.RenomeadorArquivos")
            except Exception:
                pass
            try:
                im = Image.open(ico_path).resize((32, 32), Image.Resampling.LANCZOS)
                tk_im = ImageTk.PhotoImage(im)
                self.iconphoto(True, tk_im)
                self._icon_ref = tk_im
            except Exception:
                pass
        elif os.path.exists(png_path):
            try:
                tk_im = tk.PhotoImage(file=png_path)
                self.iconphoto(True, tk_im)
                self._icon_ref = tk_im
            except Exception:
                pass

    # ---------- Estilo ----------
    def _configurar_estilo(self) -> None:
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.cores = {
            "verde_principal": "#009475",
            "verde_escuro": "#002927",
            "verde_fundo": "#F0FFFC",
            "branco": "#FFFFFF",
            "cinza_claro": "#F8FFFE",
            "cinza_medio": "#A0B5B1",
            "cinza_escuro": "#1A1A1A",
            "accent": "#00C896",
            "sombra": "#E0F5F1",
        }

        font_title = ("Segoe UI Black", 16, "bold")
        font_subtitle = ("Segoe UI Semibold", 12, "bold")
        font_text = ("Segoe UI", 11)
        font_button = ("Segoe UI Semibold", 11, "bold")

        self.style.configure("Main.TFrame", background=self.cores["verde_fundo"])
        self.style.configure(
            "Logo.TLabel",
            font=("Arial Black", 24, "bold"),
            foreground=self.cores["verde_principal"],
            background=self.cores["verde_fundo"],
        )
        self.style.configure("Title.TLabel", font=font_title, foreground=self.cores["verde_escuro"], background=self.cores["verde_fundo"])
        self.style.configure("Subtitle.TLabel", font=font_subtitle, foreground=self.cores["verde_principal"], background=self.cores["verde_fundo"])
        self.style.configure("Normal.TLabel", font=font_text, foreground=self.cores["cinza_escuro"], background=self.cores["verde_fundo"])

        self.style.configure(
            "Premium.TCombobox",
            fieldbackground=self.cores["branco"],
            borderwidth=3,
            relief="solid",
            font=font_text,
            padding=8,
        )
        self.style.map("Premium.TCombobox", fieldbackground=[("focus", self.cores["cinza_claro"]), ("readonly", self.cores["branco"])])

        self.style.configure(
            "Premium.TEntry",
            fieldbackground=self.cores["branco"],
            borderwidth=3,
            relief="solid",
            font=font_text,
            padding=8,
        )
        self.style.map("Premium.TEntry", fieldbackground=[("focus", self.cores["cinza_claro"])])

        self.style.configure(
            "Primary.TButton",
            font=font_button,
            foreground=self.cores["branco"],
            background=self.cores["verde_principal"],
            borderwidth=0,
            relief="flat",
            padding=(20, 12),
        )
        self.style.map("Primary.TButton", background=[("active", self.cores["accent"]), ("pressed", self.cores["verde_escuro"])])

        self.style.configure(
            "Secondary.TButton",
            font=font_text,
            foreground=self.cores["verde_principal"],
            background=self.cores["branco"],
            borderwidth=2,
            relief="solid",
            padding=(15, 8),
        )
        self.style.map("Secondary.TButton", background=[("active", self.cores["cinza_claro"]), ("pressed", self.cores["sombra"])])

        self.style.configure(
            "Help.TButton",
            font=("Segoe UI", 9),
            foreground=self.cores["verde_principal"],
            background=self.cores["branco"],
            borderwidth=1,
            relief="solid",
            padding=(2, 0),
        )
        self.style.map("Help.TButton", background=[("active", self.cores["cinza_claro"]), ("pressed", self.cores["sombra"])])

        self.style.configure("Preview.TFrame", background=self.cores["branco"], relief="solid", borderwidth=3)

    # ---------- UI ----------
    def _criar_interface(self) -> None:
        self.frame_principal = ttk.Frame(self, style="Main.TFrame", padding="15")
        self.frame_principal.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.frame_principal.columnconfigure(1, weight=1)

        self._criar_cabecalho()
        self._criar_secao_data()
        self._criar_secao_setor()
        self._criar_secao_funcionario()
        self._criar_secao_documento()
        self._criar_secao_versao()
        self._criar_secao_evento()
        self._criar_secao_arquivo()
        self._criar_secao_preview()
        self._criar_secao_botoes()

    def _criar_cabecalho(self) -> None:
        header = ttk.Frame(self.frame_principal, style="Main.TFrame")
        header.grid(row=0, column=0, columnspan=2, pady=(0, 20), sticky="ew")
        header.columnconfigure(1, weight=1)

        logo_widget = ttk.Label(header, image=self.logo_image, style="Logo.TLabel") if self.logo_image else ttk.Label(header, text="WF", style="Logo.TLabel")
        logo_widget.grid(row=0, column=0, rowspan=2, padx=(0, 15), sticky="w")

        titles = ttk.Frame(header, style="Main.TFrame")
        titles.grid(row=0, column=1, sticky="ew")
        ttk.Label(titles, text="Renomeador de Arquivos", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(titles, text="Instituto Walkyria Fernandes", style="Subtitle.TLabel").grid(row=1, column=0, sticky="w")

        self.btn_ajuda = ttk.Button(header, text="?", command=self._mostrar_tutorial, style="Help.TButton", width=1, cursor="hand2", takefocus=False)
        header.update_idletasks()
        self.btn_ajuda.place(relx=0.985, rely=0.10, anchor="ne")

    def _criar_secao_data(self) -> None:
        ttk.Label(self.frame_principal, text="📅 Data:", style="Normal.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Label(
            self.frame_principal,
            text=self._obter_data_atual(),
            font=("Consolas", 10, "bold"),
            foreground=self.cores["verde_principal"],
            background=self.cores["verde_fundo"],
        ).grid(row=1, column=1, sticky="w", pady=3)

    def _criar_secao_setor(self) -> None:
        ttk.Label(self.frame_principal, text="🏢 Setor:", style="Normal.TLabel").grid(row=2, column=0, sticky="w", pady=3)
        self.combo_setor = ttk.Combobox(self.frame_principal, textvariable=self.setor_selecionado, values=list(self.setores.keys()), state="readonly", style="Premium.TCombobox", width=25)
        self.combo_setor.grid(row=2, column=1, sticky="ew", pady=3)
        self.combo_setor.bind("<<ComboboxSelected>>", self._atualizar_preview)

    def _criar_secao_funcionario(self) -> None:
        ttk.Label(self.frame_principal, text="👤 Funcionário:", style="Normal.TLabel").grid(row=3, column=0, sticky="w", pady=3)
        self.combo_funcionario = ttk.Combobox(self.frame_principal, textvariable=self.funcionario_selecionado, values=self.funcionarios, state="readonly", style="Premium.TCombobox", width=25)
        self.combo_funcionario.grid(row=3, column=1, sticky="ew", pady=3)
        self.combo_funcionario.bind("<<ComboboxSelected>>", self._atualizar_preview)

    def _criar_secao_evento(self) -> None:
        ttk.Label(self.frame_principal, text="🎉 Evento:", style="Normal.TLabel").grid(row=4, column=0, sticky="w", pady=3)
        self.combo_evento = ttk.Combobox(self.frame_principal, textvariable=self.evento_selecionado, values=list(self.eventos.keys()), state="readonly", style="Premium.TCombobox", width=25)
        self.combo_evento.grid(row=4, column=1, sticky="ew", pady=4)
        self.combo_evento.bind("<<ComboboxSelected>>", self._atualizar_preview)

    def _criar_secao_documento(self) -> None:
        ttk.Label(self.frame_principal, text="📄 Documento:", style="Normal.TLabel").grid(row=5, column=0, sticky="w", pady=3)
        self.entry_documento = ttk.Entry(self.frame_principal, textvariable=self.nome_documento, style="Premium.TEntry")
        self.entry_documento.grid(row=5, column=1, sticky="ew", pady=3)
        self._configurar_placeholder(self.entry_documento, "Ex: ADITIVO CONTRATUAL")
        self.entry_documento.bind("<KeyRelease>", self._atualizar_preview)

    def _criar_secao_versao(self) -> None:
        ttk.Label(self.frame_principal, text="🔢 Versão:", style="Normal.TLabel").grid(row=6, column=0, sticky="w", pady=3)
        self.entry_versao = ttk.Entry(self.frame_principal, textvariable=self.versao_arquivo, style="Premium.TEntry")
        self.entry_versao.grid(row=6, column=1, sticky="ew", pady=3)
        self._configurar_placeholder(self.entry_versao, "Ex: 1")
        self.entry_versao.bind("<KeyRelease>", self._validar_versao)
        self.entry_versao.bind("<FocusOut>", self._atualizar_preview)

    def _criar_secao_arquivo(self) -> None:
        ttk.Label(self.frame_principal, text="📁 Arquivo:", style="Normal.TLabel").grid(row=7, column=0, sticky="w", pady=3)
        ttk.Button(self.frame_principal, text="📂 Selecionar", command=self._selecionar_arquivos, style="Primary.TButton").grid(row=7, column=1, sticky="ew", pady=3)

    def _criar_secao_preview(self) -> None:
        ttk.Label(self.frame_principal, text="👁️ Preview:", style="Normal.TLabel").grid(row=9, column=0, columnspan=2, sticky="w", pady=(15, 3))
        self.frame_preview = ttk.Frame(self.frame_principal, style="Preview.TFrame", padding="8")
        self.frame_preview.grid(row=10, column=0, columnspan=2, sticky="ew", pady=3)
        self.frame_preview.columnconfigure(0, weight=1)

        self.label_preview = ttk.Label(self.frame_preview, text="Complete os campos para ver o preview", font=("Consolas", 9, "bold"), foreground=self.cores["cinza_medio"], background=self.cores["branco"], wraplength=480)
        self.label_preview.grid(row=0, column=0, sticky="ew")

        self.label_arquivo = ttk.Label(self.frame_principal, text="Nenhum arquivo selecionado", style="Normal.TLabel", foreground=self.cores["cinza_medio"], wraplength=400)
        self.label_arquivo.grid(row=8, column=0, columnspan=2, sticky="w", pady=3)

    def _criar_secao_botoes(self) -> None:
        ttk.Button(self.frame_principal, text="✨ RENOMEAR ARQUIVO", command=self._renomear_arquivo, style="Primary.TButton").grid(row=11, column=0, columnspan=2, pady=(20, 8), sticky="ew")
        ttk.Button(self.frame_principal, text="🗑️ Limpar", command=self._limpar_campos, style="Secondary.TButton").grid(row=12, column=0, columnspan=2, pady=3, sticky="ew")

    # ---------- Helpers ----------
    def _configurar_placeholder(self, entry: ttk.Entry, placeholder: str) -> None:
        def on_focus_in(_):
            if entry.get() == placeholder:
                entry.delete(0, tk.END)
                entry.configure(foreground=self.cores["cinza_escuro"])

        def on_focus_out(_):
            if not entry.get():
                entry.insert(0, placeholder)
                entry.configure(foreground=self.cores["cinza_medio"])

        entry.insert(0, placeholder)
        entry.configure(foreground=self.cores["cinza_medio"])
        entry.bind("<FocusIn>", on_focus_in)
        entry.bind("<FocusOut>", on_focus_out)

    def _validar_versao(self, _=None) -> None:
        v = self.versao_arquivo.get()
        if v and v != "Ex: 1":
            self.versao_arquivo.set("".join(c for c in v if c.isdigit()))
        self._atualizar_preview()

    def _obter_data_atual(self) -> str:
        return datetime.date.today().strftime("%Y%m%d")

    # --- Tutorial ---
    def _tutorial_texto(self) -> str:
        return (
            "Bem-vindo!\n\n"
            "Este programa renomeia seus arquivos de forma clara e padronizada.\n\n"
            "O que você precisa saber:\n"
            "• Todos os campos são OPCIONAIS — use só o que fizer sentido para você.\n"
            "• Você pode selecionar UM ou VÁRIOS arquivos.\n"
            "  - Windows: segure Ctrl (múltiplos) ou Shift (intervalo).\n"
            "  - Mac: segure Command (múltiplos) ou Shift (intervalo).\n"
            "• O nome sugerido aparece em “Preview”.\n"
            "• Os arquivos são renomeados na MESMA pasta.\n"
            "• Em lote, a numeração será: -001, -002, -003, ...\n"
            "  (se um número já existir, o app usa o próximo livre apenas para aquele arquivo).\n"
            "• O campo Documento fica MAIÚSCULO e sem espaços (não usa hífens).\n\n"
            "Passo a passo:\n"
            "1) Clique em “📂 Selecionar” e escolha o(s) arquivo(s).\n"
            "2) (Opcional) Preencha Setor, Evento, Funcionário, Documento e Versão.\n"
            "3) Confira o “Preview”.\n"
            "4) Clique em “✨ RENOMEAR ARQUIVO”.\n\n"
            "Dica: use “🗑️ Limpar” para recomeçar quando quiser."
        )

    def _mostrar_tutorial(self) -> None:
        messagebox.showinfo("Guia rápido", self._tutorial_texto())

    def _mostrar_tutorial_se_necessario(self) -> None:
        if self.tutorial_v1_shown:
            return
        try:
            messagebox.showinfo("Guia rápido (primeira vez)", self._tutorial_texto())
        finally:
            self.tutorial_v1_shown = True
            self._salvar_configuracoes()

    # --- Seleção de arquivos ---
    def _selecionar_arquivos(self) -> None:
        tipos = [
            ("Todos os arquivos", "*.*"),
            ("Documentos PDF", "*.pdf"),
            ("Documentos Word", "*.docx"),
            ("Documentos Excel", "*.xlsx"),
            ("Imagens", "*.png *.jpg *.jpeg"),
        ]
        paths = filedialog.askopenfilenames(
            title="Selecionar arquivo(s) para renomear",
            initialdir=self.ultimo_diretorio,
            filetypes=tipos,
        )
        if not paths:
            return

        self.arquivos_selecionados = sorted(paths, key=lambda p: os.path.basename(p).lower())
        primeiro = self.arquivos_selecionados[0]
        self.caminho_arquivo.set(primeiro)
        self.ultimo_diretorio = os.path.dirname(primeiro)

        if len(self.arquivos_selecionados) == 1:
            nome = os.path.basename(primeiro)
            if len(nome) > 45:
                nome = nome[:42] + "..."
            self.label_arquivo.config(text=f"📎 {nome}", foreground=self.cores["verde_principal"])
        else:
            self.label_arquivo.config(text=f"📎 {len(self.arquivos_selecionados)} arquivos selecionados", foreground=self.cores["verde_principal"])
        self._atualizar_preview()

    # --- Sanitização ---
    @staticmethod
    def _ascii_no_accents(s: str) -> str:
        return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")

    @staticmethod
    def _sanitize_component(texto: str, *, upper: bool = False) -> str:
        if not texto:
            return ""
        s = texto.strip()
        if not s:
            return ""
        if upper:
            s = s.upper()
        s = RenomeadorArquivosApp._ascii_no_accents(s)
        s = s.replace(" ", "-")
        s = re.sub(r"[^A-Za-z0-9\-_]+", "-", s)
        s = re.sub(r"-{2,}", "-", s)
        return s.strip("-_.")

    @staticmethod
    def _sanitize_document(texto: str) -> str:
        if not texto:
            return ""
        s = texto.strip()
        if not s:
            return ""
        s = RenomeadorArquivosApp._ascii_no_accents(s.upper())
        s = s.replace(" ", "")
        return re.sub(r"[^A-Za-z0-9]", "", s)

    # --- Numeração ---
    def _num_fmt(self, n: int) -> str:
        return f"{n:03d}"

    # --- Montagem do nome ---
    def _gerar_nome_final_para(self, arquivo: str, sufixo_num: Optional[int] = None) -> Optional[str]:
        if not arquivo:
            return None

        data = self._obter_data_atual()
        setor = self.setor_selecionado.get()
        evento = self.evento_selecionado.get()
        funcionario = self.funcionario_selecionado.get()
        documento = self.nome_documento.get().strip()
        versao = self.versao_arquivo.get().strip()
        ext = os.path.splitext(arquivo)[1]

        partes: List[str] = [self._sanitize_component(data)]
        if setor in self.setores:
            partes.append(self._sanitize_component(self.setores[setor]))
        if evento in self.eventos:
            partes.append(self._sanitize_component(self.eventos[evento]))
        if funcionario:
            partes.append(self._sanitize_component(funcionario))
        if documento and documento != "Ex: ADITIVO CONTRATUAL":
            partes.append(self._sanitize_document(documento))
        if versao and versao != "Ex: 1" and versao.isdigit():
            partes.append(f"v{int(versao):02d}")

        base = "-".join(p for p in partes if p)
        if not base:
            return None

        if sufixo_num is not None:
            base = f"{base}-{self._num_fmt(sufixo_num)}"

        nome = f"{base}{ext}"
        diretorio = os.path.dirname(arquivo)
        return self._encurtar_se_preciso(diretorio, nome)

    def _gerar_nome_final(self) -> Optional[str]:
        arquivo = self.caminho_arquivo.get()
        if not arquivo:
            return None
        sufixo_num = 1 if len(self.arquivos_selecionados) > 1 else None
        return self._gerar_nome_final_para(arquivo, sufixo_num=sufixo_num)

    def _encurtar_se_preciso(self, diretorio: str, filename: str) -> str:
        full = os.path.join(diretorio, filename)
        max_full = 259 if os.name == "nt" else 4096
        name, ext = os.path.splitext(os.path.basename(filename))
        max_name = 255 - len(ext)
        if len(os.path.basename(filename)) > 255 or len(full) > max_full:
            reserva_full = 5 if os.name == "nt" else 0
            allowance_full = max_full - len(diretorio) - 1 - len(ext) - reserva_full
            allowance = max(1, min(max_name, allowance_full))
            if len(name) > allowance:
                name = name[:allowance]
        return f"{name}{ext}"

    def _atualizar_preview(self, _=None) -> None:
        nome = self._gerar_nome_final()
        if nome:
            self.label_preview.config(text=nome, foreground=self.cores["verde_principal"])
        else:
            self.label_preview.config(text="Complete os campos para ver o preview", foreground=self.cores["cinza_medio"])

    def _validar_campos(self) -> bool:
        if not self.caminho_arquivo.get() and not self.arquivos_selecionados:
            messagebox.showerror("Erro", "Selecione pelo menos um arquivo.")
            return False
        return True

    def _renomear_arquivo(self) -> None:
        if not self._validar_campos():
            return
        try:
            self._salvar_configuracoes()

            if len(self.arquivos_selecionados) > 1:
                total = len(self.arquivos_selecionados)
                ok = 0
                erros: List[str] = []

                for idx, origem in enumerate(self.arquivos_selecionados, start=1):
                    nome_final = self._gerar_nome_final_para(origem, sufixo_num=idx)
                    if not nome_final:
                        erros.append(os.path.basename(origem))
                        continue

                    diretorio = os.path.dirname(origem)
                    destino = os.path.join(diretorio, nome_final)

                    tentativa = idx
                    while os.path.exists(destino):
                        tentativa += 1
                        nome_final = self._gerar_nome_final_para(origem, sufixo_num=tentativa)
                        destino = os.path.join(diretorio, nome_final)

                    try:
                        os.replace(origem, destino)
                        ok += 1
                    except Exception:
                        erros.append(os.path.basename(origem))

                msg = f"Renomeados: {ok}/{total} arquivo(s)."
                if erros:
                    msg += "\nNão foi possível renomear:\n- " + "\n- ".join(erros[:10])
                    if len(erros) > 10:
                        msg += f"\n... (+{len(erros)-10})"
                messagebox.showinfo("Concluído ✅", msg)
                self._limpar_campos()
                return

            origem = self.caminho_arquivo.get()
            nome_final = self._gerar_nome_final_para(origem, sufixo_num=None)
            if not nome_final:
                messagebox.showerror("Erro", "Não foi possível gerar o nome final.")
                return

            destino = os.path.join(os.path.dirname(origem), self._encurtar_se_preciso(os.path.dirname(origem), nome_final))

            if os.path.exists(destino) and not messagebox.askyesno(
                "Arquivo Existe", f"O arquivo '{os.path.basename(destino)}' já existe.\nSubstituir?"
            ):
                return

            os.replace(origem, destino)
            messagebox.showinfo("Sucesso! ✅", f"Arquivo renomeado!\n\n{os.path.basename(destino)}")
            self._limpar_campos()

        except Exception as e:
            messagebox.showerror("Erro ❌", f"Erro ao renomear:\n{e}")

    def _limpar_campos(self) -> None:
        self.setor_selecionado.set("")
        self.evento_selecionado.set("")
        self.funcionario_selecionado.set("")
        self.nome_documento.set("Ex: ADITIVO CONTRATUAL")
        self.versao_arquivo.set("Ex: 1")
        self.caminho_arquivo.set("")
        self.arquivos_selecionados.clear()
        self.entry_documento.configure(foreground=self.cores["cinza_medio"])
        self.entry_versao.configure(foreground=self.cores["cinza_medio"])
        self.label_arquivo.config(text="Nenhum arquivo selecionado", foreground=self.cores["cinza_medio"])
        self._atualizar_preview()

    # ---------- Persistência ----------
    def _carregar_configuracoes(self) -> None:
        try:
            if not os.path.exists(self.config_file):
                return
            with open(self.config_file, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            s = cfg.get("ultimo_setor")
            e = cfg.get("ultimo_evento")
            fcx = cfg.get("ultimo_funcionario")
            ult_dir = cfg.get("ultimo_diretorio")
            self.tutorial_v1_shown = bool(cfg.get("tutorial_v1_shown", False))

            if s in self.setores:
                self.setor_selecionado.set(s)
            if e in self.eventos:
                self.evento_selecionado.set(e)
            if fcx in self.funcionarios:
                self.funcionario_selecionado.set(fcx)
            if ult_dir and os.path.isdir(ult_dir):
                self.ultimo_diretorio = ult_dir
        except Exception:
            pass

    def _salvar_configuracoes(self) -> None:
        try:
            cfg = {}
            if os.path.exists(self.config_file):
                try:
                    with open(self.config_file, "r", encoding="utf-8") as f:
                        cfg = json.load(f) or {}
                except Exception:
                    cfg = {}
            cfg.update(
                {
                    "ultimo_setor": self.setor_selecionado.get(),
                    "ultimo_evento": self.evento_selecionado.get(),
                    "ultimo_funcionario": self.funcionario_selecionado.get(),
                    "ultimo_diretorio": self.ultimo_diretorio,
                    "tutorial_v1_shown": self.tutorial_v1_shown,
                    "data_ultima_utilizacao": datetime.datetime.now().isoformat(),
                }
            )
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2, ensure_ascii=False)
        except Exception:
            pass


def main():
    app = RenomeadorArquivosApp()
    app.mainloop()


if __name__ == "__main__":
    main()
