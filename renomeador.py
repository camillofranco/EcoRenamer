import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.ttk import Treeview
import openpyxl
from PIL import Image, ImageOps
import fitz  # PyMuPDF para juntar PDFs
import threading
import math
import platform
import zipfile
import shutil
import subprocess

import webbrowser
import json
import urllib.request
import sys
from concurrent.futures import ThreadPoolExecutor
from tkinter import font as tkfont
from PIL import ImageTk

VERSION = "1.3.0"
UPDATE_URL = "https://raw.githubusercontent.com/camillofranco/EcoRenamer/main/version.json"
REFS_URL = "https://github.com/camillofranco/EcoRenamer/releases"
 
def resource_path(relative_path):
    """Obtém o caminho real para arquivos embutidos no binário (.exe ou .app)"""
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Suporte cross-platform para .HEIC (iPhone)
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass # Falha silenciosa caso nao exista, ignora HEIC mas salva o resto

class ToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"EcoRenamer v{VERSION}")
        self.root.geometry("850x750")
        
        # Variáveis Imagens
        self.img_folder = tk.StringVar()
        self.excel_file = tk.StringVar()
        self.digits = tk.IntVar(value=2)
        self.start_number = tk.IntVar(value=1)
        self.compress_var = tk.BooleanVar(value=True)
        self.sort_order = tk.StringVar(value="Decrescente (Z-A)")
        self.mapping = []
        
        # Variáveis PDF
        self.pdf_folder = tk.StringVar()
        self.pdf_output_name = tk.StringVar(value="Documento_Unificado.pdf")
        self.pdf_sort_order = tk.StringVar(value="Crescente (A-Z)")
        
        # Configurações de Design (Paleta EcoWave)
        self.colors = {
            "primary": "#2E7D32",     # Verde EcoWave
            "secondary": "#4527A0",   # Roxo EcoWave
            "bg": "#f5f5f5",          # Cinza Claro Fundo
            "text": "#333333",        # Texto Escuro
            "white": "#ffffff",       # Branco
            "accent": "#1B5E20",      # Verde Escuro (hovers)
            "purple_light": "#5E35B1" # Roxo Claro (hovers)
        }
        
        self.setup_styles()
        self.create_widgets()
        
    def setup_styles(self):
        style = ttk.Style()
        # No Mac, alguns elementos de estilo do ttk são limitados, então focamos no que funciona
        style.theme_use("clam") # 'clam' permite mais customização de cores que o 'aqua'
        
        style.configure("TNotebook", background=self.colors["bg"])
        style.configure("TNotebook.Tab", padding=[15, 5], font=("", 11, "bold"))
        
        style.configure("TFrame", background=self.colors["bg"])
        style.configure("TLabel", background=self.colors["bg"], foreground=self.colors["text"], font=("", 11))
        
        # Botões Customizados
        style.configure("Primary.TButton", font=("", 11, "bold"), padding=10, background=self.colors["primary"], foreground="white")
        style.map("Primary.TButton", background=[("active", self.colors["accent"])])
        
        style.configure("Secondary.TButton", font=("", 10), padding=5, background=self.colors["secondary"], foreground="white")
        style.map("Secondary.TButton", background=[("active", self.colors["purple_light"])])
        
        style.configure("Treeview", rowheight=25, font=("", 10))
        style.configure("Treeview.Heading", font=("", 10, "bold"))
        
        # Progressbar
        style.configure("Eco.Horizontal.TProgressbar", thickness=15, troughcolor=self.colors["white"], background=self.colors["primary"])

    def create_widgets(self):
        # Frame de Logo e Cabeçalho
        frame_header = tk.Frame(self.root, bg=self.colors["white"], height=100)
        frame_header.pack(side="top", fill="x")
        frame_header.pack_propagate(False)
        
        try:
            # Tenta carregar o logo embutido
            logo_img = Image.open(resource_path("logo_ecowave.png"))
            # Redimensiona mantendo proporção (ex: altura 60px)
            aspect = logo_img.width / logo_img.height
            logo_img = logo_img.resize((int(60 * aspect), 60), Image.Resampling.LANCZOS)
            self.logo_tk = ImageTk.PhotoImage(logo_img)
            
            lbl_logo = tk.Label(frame_header, image=self.logo_tk, bg=self.colors["white"])
            lbl_logo.pack(side="left", padx=20, pady=10)
        except:
            # Fallback se o logo sumir
            tk.Label(frame_header, text="eco", font=("Arial", 32, "bold"), fg=self.colors["primary"], bg=self.colors["white"]).pack(side="left", padx=(20, 0))
            tk.Label(frame_header, text="wave", font=("Arial", 32, "bold"), fg=self.colors["secondary"], bg=self.colors["white"]).pack(side="left")
        
        lbl_title = tk.Label(frame_header, text="Renomeador Profissional", font=("", 16, "bold"), fg=self.colors["text"], bg=self.colors["white"])
        lbl_title.pack(side="right", padx=20)

        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)
        
        frame_img = ttk.Frame(notebook)
        notebook.add(frame_img, text="📸 Imagens (Renomear e Comprimir)")
        
        frame_pdf = ttk.Frame(notebook)
        notebook.add(frame_pdf, text="📄 PDFs (Juntar e Comprimir)")
        
        self.create_img_widgets(frame_img)
        self.create_pdf_widgets(frame_pdf)
        
        # Barra de Status inferior com Versão e Atualização
        frame_status = ttk.Frame(self.root, padding="5")
        frame_status.pack(side="bottom", fill="x")
        
        ttk.Label(frame_status, text=f"Versão {VERSION}", font=("", 9)).pack(side="left", padx=10)
        ttk.Button(frame_status, text="Verificar Atualização", command=self.check_for_updates, style="Small.TButton").pack(side="right", padx=10)
        
        # Estilo para o botão menor
        style = ttk.Style()
        style.configure("Small.TButton", font=("", 8))

    def create_img_widgets(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        
        frame_top = ttk.Frame(parent, padding="15")
        frame_top.grid(row=0, column=0, sticky="ew")
        frame_top.columnconfigure(1, weight=1)
        
        ttk.Label(frame_top, text="Pasta de Imagens:", font=("", 12)).grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frame_top, textvariable=self.img_folder, state="readonly", font=("", 12)).grid(row=0, column=1, sticky="ew", padx=10, pady=8)
        ttk.Button(frame_top, text="Procurar...", command=self.select_img_folder).grid(row=0, column=2, pady=8)
        
        ttk.Label(frame_top, text="Planilha Excel (Opcional):", font=("", 12)).grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame_top, textvariable=self.excel_file, state="readonly", font=("", 12)).grid(row=1, column=1, sticky="ew", padx=10, pady=8)
        
        frame_botoes_excel = ttk.Frame(frame_top)
        frame_botoes_excel.grid(row=1, column=2, sticky="ew")
        ttk.Button(frame_botoes_excel, text="Procurar", command=self.select_excel, width=9).pack(side="left", padx=2)
        ttk.Button(frame_botoes_excel, text="Limpar", command=lambda: self.excel_file.set(""), width=7).pack(side="left")
        
        frame_configs = ttk.Frame(frame_top)
        frame_configs.grid(row=2, column=0, columnspan=3, sticky="ew", pady=8)
        
        ttk.Label(frame_configs, text="Dígitos (ex: 2 = 01):", font=("", 12)).pack(side="left")
        ttk.Spinbox(frame_configs, from_=1, to=10, textvariable=self.digits, width=4, font=("", 12)).pack(side="left", padx=(5, 15))
        
        ttk.Label(frame_configs, text="Nº Inicial (sem Excel):", font=("", 12)).pack(side="left")
        ttk.Spinbox(frame_configs, from_=1, to=99999, textvariable=self.start_number, width=6, font=("", 12)).pack(side="left", padx=(5, 15))
        
        ttk.Label(frame_configs, text="(Deixe o Excel em branco para usar nº em sequência!)", font=("", 10, "italic"), foreground="gray").pack(side="left")
        
        ttk.Label(frame_top, text="Ordem Temporária:", font=("", 12)).grid(row=3, column=0, sticky="w", pady=8)
        opcoes_ordem = ["Decrescente (Z-A)", "Crescente (A-Z)"]
        cb_ordem = ttk.Combobox(frame_top, textvariable=self.sort_order, values=opcoes_ordem, state="readonly", font=("", 12), width=18)
        cb_ordem.grid(row=3, column=1, sticky="w", padx=10, pady=8)
        
        ttk.Checkbutton(frame_top, text="Comprimir Fotos (Funciona p/ Windows e Mac)", variable=self.compress_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=8)
        
        btn_load = ttk.Button(frame_top, text="1. Carregar Prévia do Mapeamento", command=self.load_data, style="Primary.TButton")
        btn_load.grid(row=5, column=0, columnspan=2, pady=15, sticky="ew")
        
        btn_clear_table = ttk.Button(frame_top, text="Limpar Tabela", command=self.reset_preview)
        btn_clear_table.grid(row=5, column=2, pady=15, padx=(5, 0), sticky="ew")
        
        frame_mid = ttk.Frame(parent, padding="15")
        frame_mid.grid(row=1, column=0, sticky="nsew")
        parent.rowconfigure(1, weight=1)
        frame_mid.columnconfigure(0, weight=1)
        frame_mid.rowconfigure(0, weight=1)
        
        columns = ("pos", "orig", "novo", "tam_orig", "tam_est")
        self.tree = Treeview(frame_mid, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("pos", text="Posição")
        self.tree.heading("orig", text="Nome Original")
        self.tree.heading("novo", text="Novo Nome Final")
        self.tree.heading("tam_orig", text="Tam. Original")
        self.tree.heading("tam_est", text="Tam. Est. (KB)")
        
        self.tree.column("pos", width=60, anchor="center")
        self.tree.column("orig", width=250, anchor="w")
        self.tree.column("novo", width=250, anchor="w")
        self.tree.column("tam_orig", width=100, anchor="center")
        self.tree.column("tam_est", width=100, anchor="center")
        
        scrollbar = ttk.Scrollbar(frame_mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Bindings para Drag and Drop
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_drop)
        self.drag_data = {"item": None}
        
        frame_bot = ttk.Frame(parent, padding="15")
        frame_bot.grid(row=2, column=0, sticky="ew")
        frame_bot.columnconfigure(0, weight=1)
        
        self.btn_rename = ttk.Button(frame_bot, text="2. Renomear Arquivos", command=self.rename_files, state="disabled", style="Primary.TButton")
        self.btn_rename.grid(row=0, column=0, pady=5, sticky="ew")
        
        # Novo sistema de progresso com texto
        self.frame_progress = ttk.Frame(frame_bot)
        self.frame_progress.grid(row=1, column=0, sticky="ew")
        self.frame_progress.grid_remove()
        
        self.lbl_status = ttk.Label(self.frame_progress, text="Aguardando...", font=("", 10))
        self.lbl_status.pack(side="top", anchor="w")
        
        self.progress = ttk.Progressbar(self.frame_progress, orient="horizontal", length=100, mode="determinate", style="Eco.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(2, 0))
        
        self.lbl_perc = ttk.Label(self.frame_progress, text="0%", font=("", 9, "bold"))
        self.lbl_perc.place(relx=0.5, rely=0.6, anchor="center")

    def create_pdf_widgets(self, parent):
        parent.columnconfigure(0, weight=1)
        
        frame_top = ttk.Frame(parent, padding="15")
        frame_top.grid(row=0, column=0, sticky="ew")
        frame_top.columnconfigure(1, weight=1)
        
        ttk.Label(frame_top, text="Pasta com PDFs:", font=("", 12)).grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frame_top, textvariable=self.pdf_folder, state="readonly", font=("", 12)).grid(row=0, column=1, sticky="ew", padx=10, pady=8)
        ttk.Button(frame_top, text="Procurar...", command=self.select_pdf_folder).grid(row=0, column=2, pady=8)
        
        ttk.Label(frame_top, text="Nome do Arquivo Final:", font=("", 12)).grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame_top, textvariable=self.pdf_output_name, font=("", 12)).grid(row=1, column=1, sticky="ew", padx=10, pady=8)
        
        ttk.Label(frame_top, text="Ordem de Leitura:", font=("", 12)).grid(row=2, column=0, sticky="w", pady=8)
        opcoes_ordem_pdf = ["Crescente (A-Z)", "Decrescente (Z-A)"]
        cb_ordem_pdf = ttk.Combobox(frame_top, textvariable=self.pdf_sort_order, values=opcoes_ordem_pdf, state="readonly", font=("", 12), width=18)
        cb_ordem_pdf.grid(row=2, column=1, sticky="w", padx=10, pady=8)
        
        btn_merge = ttk.Button(parent, text="Juntar e Comprimir todos os PDFs", command=self.merge_pdfs)
        btn_merge.grid(row=1, column=0, pady=30, padx=20, sticky="ew", ipady=10)
        
        # Informativo
        info_text = ("Selecione uma pasta que tenha os seus documentos em PDF.\n"
                     "Eles serão lidos na ordem escolhida, unificados num único arquivo\n"
                     "e comprimidos automaticamente (livrando-se de dados ocultos inúteis do PDF).")
        ttk.Label(parent, text=info_text, justify="center", font=("", 11, "italic")).grid(row=2, column=0, pady=20)

    # ------------------ UTILITÁRIOS ------------------
    def check_for_updates(self):
        try:
            with urllib.request.urlopen(UPDATE_URL, timeout=5) as response:
                data = json.loads(response.read().decode())
                remote_version = data.get("version", VERSION)
                changelog = data.get("changelog", "Melhorias gerais.")
                
                if remote_version > VERSION:
                    # Encontrar URL correta para o SO atual
                    os_name = platform.system()
                    if os_name == "Darwin":
                        download_url = data.get("download_url_mac")
                    else:
                        download_url = data.get("download_url_win")

                    msg = f"Uma nova versão ({remote_version}) está disponível!\n\n📋 O que mudou:\n{changelog}\n\nDeseja atualizar agora automaticamente?"
                    
                    if messagebox.askyesno("Atualização Disponível", msg):
                        if download_url and "/releases/download/" in download_url:
                            # Tenta baixar e "instalar" (abrir o novo)
                            threading.Thread(target=self.run_auto_update, args=(download_url, remote_version), daemon=True).start()
                        else:
                            # Caso não tenha URL direta, abre o site
                            webbrowser.open(REFS_URL)
                else:
                    messagebox.showinfo("Atualizado", f"Você já está usando a versão mais recente ({VERSION}).")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível verificar atualizações:\n{e}")

    def run_auto_update(self, url, version):
        try:
            # Pasta de destino temporária
            home = os.path.expanduser("~")
            temp_dir = os.path.join(home, "Downloads", f"EcoRenamer_Update_{version}")
            if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            zip_path = os.path.join(temp_dir, "update.zip")
            
            # 1. Download
            self.root.after(0, lambda: messagebox.showinfo("Baixando", "A atualização está sendo baixada. Isso pode levar alguns segundos..."))
            urllib.request.urlretrieve(url, zip_path)
            
            # 2. Extração
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            os.remove(zip_path) # Limpa o zip
            
            # 3. Informar e Abrir
            msg_fin = (f"Atualização baixada com sucesso em:\n{temp_dir}\n\n"
                       "A pasta com a nova versão será aberta agora. "
                       "Você pode fechar esta versão antiga e usar a nova!")
            
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", msg_fin))
            
            # Abrir pasta/arquivo
            if platform.system() == "Darwin":
                subprocess.run(["open", temp_dir])
            else:
                os.startfile(temp_dir)
                
            # Fecha o app atual opcionalmente? 
            # Melhor deixar o usuário fechar pra ele não perder o que estava fazendo
            # mas vamos dar a dica.
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Erro no Download", f"Falha ao baixar atualização automática:\n{e}\n\nTentando abrir página manual..."))
            self.root.after(0, lambda: webbrowser.open(REFS_URL))

    def format_size(self, size_bytes):
        if size_bytes >= 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / 1024:.0f} KB"

    # ------------------ DRAG & DROP ------------------
    def on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_data["item"] = item
            self.tree.config(cursor="hand2")

    def on_drag_motion(self, event):
        pass # Poderia adicionar um indicador visual aqui

    def on_drag_drop(self, event):
        self.tree.config(cursor="")
        if not self.drag_data["item"]:
            return
            
        target_item = self.tree.identify_row(event.y)
        source_item = self.drag_data["item"]
        
        if target_item and target_item != source_item:
            source_idx = self.tree.index(source_item)
            target_idx = self.tree.index(target_item)
            
            # Reordenar self.mapping
            moved_item = self.mapping.pop(source_idx)
            self.mapping.insert(target_idx, moved_item)
            
            # Recalcular nomes e atualizar Treeview
            self.update_mapping_after_reorder()
            
        self.drag_data["item"] = None

    def update_mapping_after_reorder(self):
        # Em vistorias, a ordem do Excel (destino) é geralmente fixa por posição.
        # Se o usuário arrasta a foto, ele quer mudar QUAL foto vai para aquela posição do Excel.
        # Portanto, mantemos a lista de 'img_val' originais (que guardamos no mapping).
        # Na verdade, o mais simples é: cada posição 'i' na tabela corresponde ao 'Novo Nome' do item 'i' original.
        
        # Vamos coletar todos os 'target_base' disponíveis na ordem original
        # mas como eles podem ter sido carregados de forma diferente, 
        # o ideal é ter guardado a lista de nomes destino originais no momento do load_data.
        
        if not hasattr(self, 'original_dest_bases') or not self.original_dest_bases:
            return

        digitos = self.digits.get()
        folder = self.img_folder.get()
        compress = self.compress_var.get()

        for i, item in enumerate(self.mapping):
            if i < len(self.original_dest_bases):
                img_val_str = self.original_dest_bases[i]
                
                if img_val_str.isdigit():
                    novo_base = img_val_str.zfill(digitos)
                else:
                    novo_base = img_val_str
                
                _, ext = os.path.splitext(item['orig_name'])
                if compress:
                    item['new_name'] = f"{novo_base}.JPG"
                else:
                    item['new_name'] = f"{novo_base}{ext.upper()}"
                
                item['new_path'] = os.path.join(folder, item['new_name'])

        # Atualizar Treeview
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
            
        for idx, item in enumerate(self.mapping):
            self.tree.insert("", "end", values=(idx + 1, item['orig_name'], item['new_name'], item['size_orig_str'], item['size_est_str']))

    # ------------------ LÓGICA IMAGENS ------------------
    def select_img_folder(self):
        folder = filedialog.askdirectory(title="Selecione a pasta com as fotos")
        if folder:
            self.img_folder.set(folder)
            self.reset_preview()

    def select_excel(self):
        file = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel files", "*.xlsx")])
        if file:
            self.excel_file.set(file)
            self.reset_preview()

    def reset_preview(self):
        self.mapping = []
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.btn_rename.config(state="disabled")

    def load_data(self):
        folder = self.img_folder.get()
        excel_path = self.excel_file.get()
        
        if not folder:
            messagebox.showwarning("Aviso", "Por favor, selecione a pasta de imagens.")
            return

        valid_exts = {".jpg", ".jpeg", ".png", ".heic"}
        try:
            arquivos = os.listdir(folder)
            images = [f for f in arquivos if not f.startswith('.') and os.path.splitext(f)[1].lower() in valid_exts]
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao acessar a pasta selecionada:\n{e}")
            return
            
        if not images:
            messagebox.showwarning("Aviso", "Nenhuma imagem válida encontrada na pasta selecionada.")
            return
            
        ordem = self.sort_order.get()
        if "Decrescente" in ordem:
            images.sort(reverse=True)
            self.tree.heading("orig", text="Nome Original (Decrescente)")
        else:
            images.sort(reverse=False)
            self.tree.heading("orig", text="Nome Original (Crescente)")
        
        excel_img_values = []
        if excel_path:
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=True)
                sheet = wb.active
                
                img_col_idx = None
                for col_idx in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=1, column=col_idx).value
                    if cell_value and str(cell_value).strip().upper() == "IMG":
                        img_col_idx = col_idx
                        break
                        
                if img_col_idx is None:
                    messagebox.showerror("Erro", "Coluna 'IMG' não encontrada na primeira linha do Excel.")
                    return
                    
                for row_idx in range(2, sheet.max_row + 1):
                    if row_idx in sheet.row_dimensions and sheet.row_dimensions[row_idx].hidden:
                        continue
                    val = sheet.cell(row=row_idx, column=img_col_idx).value
                    if val is not None:
                        excel_img_values.append(val)
                        
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao abrir o arquivo Excel:\n{e}")
                return
                
            if not excel_img_values:
                messagebox.showwarning("Aviso", "Nenhum valor encontrado na coluna 'IMG' nas linhas visíveis.")
                return
                
            if len(images) != len(excel_img_values):
                msg = f"A quantidade de imagens ({len(images)}) difere do Excel ({len(excel_img_values)})."
                messagebox.showwarning("Incompatibilidade", msg)
                
            qtd_pares = min(len(images), len(excel_img_values))
        else:
            qtd_pares = len(images)
            start_no = self.start_number.get()
            excel_img_values = [str(start_no + i) for i in range(qtd_pares)]
            
        self.mapping = []
        self.original_dest_bases = [] # Salvar para reordenação rápida
        digitos = self.digits.get()
        seen_targets = set()
        duplicados = set()
        
        for i in range(qtd_pares):
            orig_name = images[i]
            img_val = excel_img_values[i]
            
            img_val_str = str(img_val).strip()
            if img_val_str.endswith(".0"):
                img_val_str = img_val_str[:-2]
                
            if img_val_str.isdigit():
                novo_base = img_val_str.zfill(digitos)
            else:
                novo_base = img_val_str
                
            self.original_dest_bases.append(img_val_str) # Guarda o valor base sem preenchimento de zeros fixo
                
            _, ext = os.path.splitext(orig_name)
            if self.compress_var.get():
                novo_nome = f"{novo_base}.JPG"
            else:
                novo_nome = f"{novo_base}{ext.upper()}"
            
            if novo_nome in seen_targets:
                duplicados.add(novo_nome)
            seen_targets.add(novo_nome)
            
            orig_path = os.path.join(folder, orig_name)
            novo_path = os.path.join(folder, novo_nome)
            
            self.mapping.append({
                'orig_name': orig_name,
                'new_name': novo_nome,
                'orig_path': orig_path,
                'new_path': novo_path
            })
            
        if duplicados:
            msg_dup = f"ATENÇÃO: Este mapeamento criará arquivos nomes idênticos: {', '.join(list(duplicados)[:5])}"
            messagebox.showwarning("Conflito de Nomes", msg_dup)
            
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for idx, item in enumerate(self.mapping):
            size_orig = os.path.getsize(item['orig_path'])
            item['size_orig_str'] = self.format_size(size_orig)
            
            # Estimativa: largura 800px x 45% qualidade costuma ficar entre 40 e 90KB
            # Usaremos 70KB como média segura para exibição
            item['size_est_str'] = "70 KB" 
            
            self.tree.insert("", "end", values=(idx + 1, item['orig_name'], item['new_name'], item['size_orig_str'], item['size_est_str']))
            
        if self.mapping:
            self.btn_rename.config(state="normal")
            
    def rename_files(self):
        if not self.mapping or self.processing:
            return
            
        resposta = messagebox.askyesno("Confirmar", f"Deseja renomear {len(self.mapping)} arquivos?")
        if not resposta:
            return
            
        self.processing = True
        self.btn_rename.config(state="disabled")
        self.frame_progress.grid()
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.mapping)
        self.lbl_perc.config(text="0%")
        self.lbl_status.config(text="Iniciando processamento paralelo...")
        
        threading.Thread(target=self.run_rename_task_parallel, daemon=True).start()

    def run_rename_task_parallel(self):
        sucessos = 0
        falhas = 0
        erros_msg = []
        
        total = len(self.mapping)
        
        # Usamos ThreadPoolExecutor para processar imagens paralelamente
        # No Mac, o limitador costuma ser CPU para compressão e I/O
        # max_workers=None usa núcleos disponíveis
        with ThreadPoolExecutor() as executor:
            # Envia tarefas
            futures = []
            for idx, item in enumerate(self.mapping):
                futures.append(executor.submit(self.process_single_image, item, idx))
                
            # Coleta resultados
            for i, future in enumerate(futures):
                res_ok, res_msg = future.result()
                if res_ok:
                    sucessos += 1
                else:
                    falhas += 1
                    erros_msg.append(res_msg)
                
                # Atualiza UI
                perc = int(((i + 1) / total) * 100)
                status_txt = f"Processando {i+1} de {total}..."
                if i < total - 1:
                    status_txt += f" (Atual: {self.mapping[i+1]['orig_name']})"
                
                self.root.after(0, lambda v=i+1, p=perc, s=status_txt: self.update_ui_progress(v, p, s))

        # Finaliza na interface
        self.root.after(0, lambda: self.finish_rename(sucessos, falhas, erros_msg))

    def update_ui_progress(self, val, perc, status):
        self.progress['value'] = val
        self.lbl_perc.config(text=f"{perc}%")
        self.lbl_status.config(text=status)

    def process_single_image(self, item, index):
        """Função executada em thread separada para cada imagem"""
        if os.path.exists(item['new_path']) and item['orig_path'] != item['new_path']:
            return False, f"{item['orig_name']} -> O destino {item['new_name']} já existe."
        
        try:
            if self.compress_var.get():
                img = Image.open(item['orig_path'])
                img = ImageOps.exif_transpose(img)
                img.thumbnail((800, 800), Image.Resampling.LANCZOS)
                img = img.convert("RGB")
                
                if item['orig_path'] == item['new_path']:
                    temp_path = item['orig_path'] + "_temp.jpg"
                    img.save(temp_path, "JPEG", optimize=True, quality=45)
                    img.close()
                    os.remove(item['orig_path'])
                    os.rename(temp_path, item['new_path'])
                else:
                    img.save(item['new_path'], "JPEG", optimize=True, quality=45)
                    img.close()
                    os.remove(item['orig_path'])
            else:
                if item['orig_path'] != item['new_path']:
                    shutil.copy2(item['orig_path'], item['new_path']) # Usamos copy2 e depois remove p/ segurança
                    os.remove(item['orig_path'])
            return True, ""
        except Exception as e:
            return False, f"{item['orig_name']}: {str(e)}"

        # Finaliza na interface
        self.root.after(0, lambda: self.finish_rename(sucessos, falhas, erros_msg))

    def finish_rename(self, sucessos, falhas, erros_msg):
        msg = f"Processo concluído!\nSucesso: {sucessos}\nFalhas: {falhas}"
        if falhas > 0:
            msg += f"\n\nErros (máximo 5):\n" + "\n".join(erros_msg[:5])
            messagebox.showwarning("Aviso", msg)
        else:
            messagebox.showinfo("Sucesso", msg)
        
        self.processing = False
        self.frame_progress.grid_remove() # Esconde barra
        self.reset_preview()

    # ------------------ LÓGICA PDFs ------------------
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
        if folder:
            self.pdf_folder.set(folder)

    def merge_pdfs(self):
        folder = self.pdf_folder.get()
        if not folder:
            messagebox.showwarning("Aviso", "Selecione a pasta com os PDFs primeiro!")
            return
            
        try:
            arquivos = os.listdir(folder)
            pdfs = [f for f in arquivos if not f.startswith('.') and f.lower().endswith('.pdf')]
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao acessar a pasta selecionada:\n{e}")
            return
            
        if not pdfs:
            messagebox.showwarning("Aviso", "Nenhum PDF encontrado na pasta selecionada.")
            return
            
        ordem = self.pdf_sort_order.get()
        if "Decrescente" in ordem:
            pdfs.sort(reverse=True)
        else:
            pdfs.sort(reverse=False) # Ordem padrão pra PDF costuma ser página 1 primeiro, então crescente
            
        output_name = self.pdf_output_name.get().strip()
        if not output_name:
            output_name = "Documento_Unificado.pdf"
            
        if not output_name.lower().endswith(".pdf"):
            output_name += ".pdf"
            
        output_path = os.path.join(folder, output_name)
        
        # Evitar loop infinito se o arquivo q estamos querendo gerar já estiver na pasta e na lista
        if output_name in pdfs:
            pdfs.remove(output_name)
            
        if not pdfs:
            messagebox.showwarning("Aviso", "Nenhum PDF para processar (além do próprio arquivo final já existente).")
            return
            
        try:
            # Cria doc PDF em branco
            doc_final = fitz.open() 
            
            for pdf_file in pdfs:
                pdf_path = os.path.join(folder, pdf_file)
                try:
                    doc_temp = fitz.open(pdf_path)
                    doc_final.insert_pdf(doc_temp)
                    doc_temp.close()
                except Exception as e:
                    print(f"Ignorando arquivo defeituoso {pdf_file}: {e}")
            
            # Salvar usando Garbage=4 e Deflate, a melhor forma de comprimir PDF no fitz (PyMuPDF)
            # Ele limpará fontes não usadas e streams duplicadas
            doc_final.save(output_path, garbage=4, deflate=True)
            doc_final.close()
            
            messagebox.showinfo("Concluído", f"Sucesso!\n{len(pdfs)} arquivos foram unificados e salvos como '{output_name}' na mesma pasta.\n\nEles foram comprimidos (dados não usados apagados e otimizados) automaticamente.")
            
        except Exception as e:
            messagebox.showerror("Erro Fatal", f"Ocorreu um erro ao processar os PDFs:\n\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ToolApp(root)
    try:
        style = ttk.Style()
        if "aqua" in style.theme_names():
            style.theme_use("aqua")
    except:
        pass
    root.mainloop()
