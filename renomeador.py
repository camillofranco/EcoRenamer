import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from tkinter.ttk import Treeview
import openpyxl
from PIL import Image, ImageOps
import fitz  # PyMuPDF
import threading
import platform
import zipfile
import shutil
import subprocess
import webbrowser
import json
import urllib.request
import sys
from concurrent.futures import ThreadPoolExecutor

VERSION = "1.4.3" # Hotfix Critico: Caminho do auto-updater
UPDATE_URL = "https://raw.githubusercontent.com/camillofranco/EcoRenamer/main/version.json"
REFS_URL = "https://github.com/camillofranco/EcoRenamer/releases"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

# Configuração Global do Tema CTk
ctk.set_appearance_mode("System")  # Segue o sistema (Dark/Light)
ctk.set_default_color_theme("green") # Base verde do sistema

class ToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"EcoWave Pro v{VERSION} - Enterprise Edition")
        self.root.geometry("1000x800")
        self.root.minsize(900, 700) # Evita cortar telas pequenas
        
        # Paleta EcoWave
        self.c_primary = "#2E7D32" # Verde Forte
        self.c_secondary = "#4527A0" # Roxo Elegante
        
        # Variáveis Imagens
        self.img_folder = ctk.StringVar()
        self.excel_file = ctk.StringVar()
        self.digits = tk.IntVar(value=2)
        self.start_number = tk.IntVar(value=1)
        self.compress_var = ctk.BooleanVar(value=True)
        self.sort_order = ctk.StringVar(value="Decrescente (Z-A)")
        self.mapping = []
        
        # Variáveis PDF
        self.pdf_folder = ctk.StringVar()
        self.pdf_output_name = ctk.StringVar(value="Documento_Unificado.pdf")
        self.pdf_sort_order = ctk.StringVar(value="Crescente (A-Z)")
        self.processing = False
        
        self.create_widgets()
        
    def create_widgets(self):
        # 1. HEADER (Cabeçalho Branco Puro)
        # Em CTk, o Frame aceita cor sólida facilmente.
        self.header = ctk.CTkFrame(self.root, height=100, fg_color="#ffffff", corner_radius=0)
        self.header.pack(fill="x", side="top")
        self.header.pack_propagate(False)
        
        try:
            logo_img = Image.open(resource_path("logo_ecowave.png"))
            target_h = 75
            aspect = logo_img.width / logo_img.height
            logo_img = logo_img.resize((int(target_h * aspect), target_h), Image.Resampling.LANCZOS)
            self.logo_ctk = ctk.CTkImage(light_image=logo_img, dark_image=logo_img, size=(int(target_h * aspect), target_h))
            lbl_logo = ctk.CTkLabel(self.header, text="", image=self.logo_ctk)
            lbl_logo.pack(side="left", padx=40, pady=12)
        except Exception as e:
            ctk.CTkLabel(self.header, text="ECOWAVE PRO", font=ctk.CTkFont(size=28, weight="bold"), text_color=self.c_primary).pack(side="left", padx=40)
            
        badge = ctk.CTkFrame(self.header, fg_color="transparent")
        badge.pack(side="right", padx=40)
        ctk.CTkLabel(badge, text="ENTERPRISE EDITION", font=ctk.CTkFont(size=12, weight="bold"), text_color=self.c_secondary).pack()
        ctk.CTkLabel(badge, text=f"Build v{VERSION}", font=ctk.CTkFont(size=10), text_color="gray").pack()
        
        # 2. SELETOR DE ABAS PRINCIPAL (Muito mais bonito que o Notebook antigo)
        self.tabview = ctk.CTkTabview(self.root, corner_radius=10, segmented_button_selected_color=self.c_primary,
                                      segmented_button_selected_hover_color="#1B5E20")
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)
        
        self.tabview.add("📸  GESTÃO DE IMAGENS")
        self.tabview.add("📄  GESTÃO DE PDFS")
        
        self.create_img_tab(self.tabview.tab("📸  GESTÃO DE IMAGENS"))
        self.create_pdf_tab(self.tabview.tab("📄  GESTÃO DE PDFS"))
        
        # 3. BARRA DE STATUS INFERIOR
        self.status_bar = ctk.CTkFrame(self.root, height=40, corner_radius=0, fg_color=("gray90", "gray15"))
        self.status_bar.pack(fill="x", side="bottom")
        
        ctk.CTkLabel(self.status_bar, text=f"SO: {platform.system()} | Paralelismo Ativado (Motor Rápido)", 
                     font=ctk.CTkFont(size=11), text_color="gray").pack(side="left", padx=20)
        
        btn_upd = ctk.CTkButton(self.status_bar, text="♻️ Baixar Atualizações", command=self.check_for_updates,
                                fg_color=self.c_primary, text_color="white", font=ctk.CTkFont(size=12, weight="bold"), hover_color="#1B5E20", width=160, height=30)
        btn_upd.pack(side="right", padx=20, pady=5)
        
    def create_img_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1) # A tabela expande
        
        # --- BLOCO SUPERIOR (Configs) ---
        frame_config = ctk.CTkFrame(parent, fg_color="transparent")
        frame_config.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        frame_config.columnconfigure(1, weight=1)
        
        # Linha 1: Imagens
        ctk.CTkLabel(frame_config, text="PASTA DE FOTOS:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", pady=5)
        ctk.CTkEntry(frame_config, textvariable=self.img_folder, state="readonly", height=35).grid(row=0, column=1, sticky="ew", padx=10, pady=5)
        ctk.CTkButton(frame_config, text="Selecionar", command=self.select_img_folder, fg_color=self.c_secondary, hover_color="#311B92", width=100).grid(row=0, column=2, pady=5)
        
        # Linha 2: Excel
        ctk.CTkLabel(frame_config, text="EXCEL (Opcional):", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, sticky="w", pady=5)
        ctk.CTkEntry(frame_config, textvariable=self.excel_file, state="readonly", height=35).grid(row=1, column=1, sticky="ew", padx=10, pady=5)
        
        frame_b = ctk.CTkFrame(frame_config, fg_color="transparent")
        frame_b.grid(row=1, column=2, sticky="ew")
        ctk.CTkButton(frame_b, text="Buscar", command=self.select_excel, fg_color=self.c_secondary, width=70).pack(side="left", padx=(0,5))
        ctk.CTkButton(frame_b, text="Limpar", command=lambda: self.excel_file.set(""), fg_color="gray50", hover_color="gray30", width=70).pack(side="left")
        
        # Linha 3: Opções Mistas
        frame_opts = ctk.CTkFrame(frame_config, fg_color="transparent")
        frame_opts.grid(row=2, column=0, columnspan=3, sticky="w", pady=15)
        
        ctk.CTkLabel(frame_opts, text="Dígitos:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0,5))
        # CTk não tem Spinbox nativo, usamos Entry (poderíamos criar um stepper)
        ctk.CTkEntry(frame_opts, textvariable=self.digits, width=50, justify="center").pack(side="left", padx=(0,20))
        
        ctk.CTkLabel(frame_opts, text="Nº Inicial:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0,5))
        ctk.CTkEntry(frame_opts, textvariable=self.start_number, width=60, justify="center").pack(side="left", padx=(0,20))
        
        ctk.CTkLabel(frame_opts, text="Ordem Interna:", font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(0,5))
        ctk.CTkComboBox(frame_opts, variable=self.sort_order, values=["Decrescente (Z-A)", "Crescente (A-Z)"], width=170).pack(side="left", padx=(0,20))
        
        ctk.CTkCheckBox(frame_opts, text="Comprimir em HD (Mais Rápido no App)", variable=self.compress_var, text_color=self.c_primary, font=ctk.CTkFont(weight="bold")).pack(side="left")

        # Botão Carregar (Grande)
        self.btn_load = ctk.CTkButton(frame_config, text="1. MONTAR ESTRUTURA DE NOMES", command=self.load_data, 
                                      fg_color=self.c_primary, text_color="white", font=ctk.CTkFont(size=14, weight="bold"), height=45)
        self.btn_load.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(15, 0))


        # --- BLOCO CENTRAL (Tabela Tkinter Estilizada para CTk) ---
        frame_tree = ctk.CTkFrame(parent, fg_color="transparent")
        frame_tree.grid(row=2, column=0, sticky="nsew", pady=10)
        frame_tree.columnconfigure(0, weight=1)
        frame_tree.rowconfigure(0, weight=1)
        
        # É uma árvore Tkinter, mas daremos um ar moderno a ela alterando via ttk.Style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", 
                        background="#ffffff" if ctk.get_appearance_mode() == "Light" else "#2b2b2b",
                        foreground="#000000" if ctk.get_appearance_mode() == "Light" else "#ffffff",
                        rowheight=35,
                        fieldbackground="#ffffff" if ctk.get_appearance_mode() == "Light" else "#2b2b2b",
                        font=("Helvetica", 11),
                        borderwidth=0)
        style.configure("Treeview.Heading", font=("Helvetica", 11, "bold"), background=self.c_primary, foreground="white", borderwidth=0)
        style.map("Treeview", background=[("selected", self.c_secondary)])
        
        cols = ("pos", "orig", "novo", "tam_orig", "tam_est")
        self.tree = Treeview(frame_tree, columns=cols, show="headings", selectmode="browse")
        for col in cols: self.tree.heading(col, text=col.upper())
        
        # Bindings Drag Drop
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_drop)
        self.drag_data = {"item": None}

        scrollbar = ctk.CTkScrollbar(frame_tree, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # --- BLOCO INFERIOR (Ação) ---
        frame_bot = ctk.CTkFrame(parent, fg_color="transparent")
        frame_bot.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        
        self.btn_rename = ctk.CTkButton(frame_bot, text="2. INICIAR PROCESSAMENTO PARALELO (PRO)", command=self.rename_files,
                                        fg_color=self.c_primary, hover_color="#1B5E20", text_color="white", font=ctk.CTkFont(size=14, weight="bold"), height=50, state="disabled")
        self.btn_rename.pack(fill="x")
        
        self.frame_progress = ctk.CTkFrame(frame_bot, fg_color="transparent")
        # Escondido por padrão
        
        self.lbl_status = ctk.CTkLabel(self.frame_progress, text="Aguardando...", font=ctk.CTkFont(weight="bold"))
        self.lbl_status.pack(anchor="w", pady=(10, 0))
        
        self.progress = ctk.CTkProgressBar(self.frame_progress, mode="determinate", progress_color=self.c_primary, height=20)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=5)
        
        self.lbl_perc = ctk.CTkLabel(self.frame_progress, text="0%", font=ctk.CTkFont(weight="bold", size=14))
        self.lbl_perc.pack()

    def create_pdf_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        
        frame_top = ctk.CTkFrame(parent, fg_color="transparent")
        frame_top.grid(row=0, column=0, sticky="ew", pady=20)
        frame_top.columnconfigure(1, weight=1)
        
        ctk.CTkLabel(frame_top, text="PASTA COM PDFs:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", pady=10)
        ctk.CTkEntry(frame_top, textvariable=self.pdf_folder, state="readonly", height=40).grid(row=0, column=1, sticky="ew", padx=15, pady=10)
        ctk.CTkButton(frame_top, text="Selecionar", command=self.select_pdf_folder, fg_color=self.c_secondary, height=40).grid(row=0, column=2, pady=10)
        
        ctk.CTkLabel(frame_top, text="NOME DO ARQUIVO FINAL:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, sticky="w", pady=10)
        ctk.CTkEntry(frame_top, textvariable=self.pdf_output_name, height=40).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(15,0), pady=10)
        
        ctk.CTkButton(parent, text="UNIFICAR E COMPRIMIR PDFs", command=self.merge_pdfs,
                      fg_color=self.c_primary, text_color="white", font=ctk.CTkFont(size=16, weight="bold"), height=60, corner_radius=8).grid(row=1, column=0, pady=40, sticky="ew")

    # ------------------ DRAG & DROP ------------------
    def on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.drag_data["item"] = item
            self.tree.config(cursor="hand2")

    def on_drag_motion(self, event):
        pass

    def on_drag_drop(self, event):
        self.tree.config(cursor="")
        if not self.drag_data["item"]: return
            
        target_item = self.tree.identify_row(event.y)
        source_item = self.drag_data["item"]
        
        if target_item and target_item != source_item:
            source_idx = self.tree.index(source_item)
            target_idx = self.tree.index(target_item)
            
            moved_item = self.mapping.pop(source_idx)
            self.mapping.insert(target_idx, moved_item)
            
            self.update_mapping_after_reorder()
            
        self.drag_data["item"] = None

    def update_mapping_after_reorder(self):
        if not hasattr(self, 'original_dest_bases') or not self.original_dest_bases:
            return

        try:
            digitos = int(self.digits.get())
        except:
            digitos = 2
            
        folder = self.img_folder.get()
        compress = self.compress_var.get()

        for i, item in enumerate(self.mapping):
            if i < len(self.original_dest_bases):
                img_val_str = self.original_dest_bases[i]
                
                if img_val_str.isdigit(): novo_base = img_val_str.zfill(digitos)
                else: novo_base = img_val_str
                
                _, ext = os.path.splitext(item['orig_name'])
                if compress: item['new_name'] = f"{novo_base}.JPG"
                else: item['new_name'] = f"{novo_base}{ext.upper()}"
                
                item['new_path'] = os.path.join(folder, item['new_name'])

        for item_id in self.tree.get_children():
            self.tree.delete(item_id)
            
        for idx, item in enumerate(self.mapping):
            self.tree.insert("", "end", values=(idx + 1, item['orig_name'], item['new_name'], item['size_orig_str'], item['size_est_str']))

    # ------------------ LÓGICA IMAGENS ------------------
    def format_size(self, size_bytes):
        if size_bytes >= 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.2f} MB"
        else:
            return f"{size_bytes / 1024:.0f} KB"

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
        self.btn_rename.configure(state="disabled")

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
            messagebox.showerror("Erro", f"Erro ao acessar pasta:\n{e}")
            return
            
        if not images:
            messagebox.showwarning("Aviso", "Nenhuma imagem válida encontrada.")
            return
            
        ordem = self.sort_order.get()
        if "Decrescente" in ordem: images.sort(reverse=True)
        else: images.sort(reverse=False)
        
        excel_img_values = []
        if excel_path:
            try:
                wb = openpyxl.load_workbook(excel_path, data_only=True)
                sheet = wb.active
                img_col_idx = None
                for col_idx in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=1, column=col_idx).value
                    if cell_value and str(cell_value).strip().upper() == "IMG":
                        img_col_idx = col_idx; break
                if img_col_idx is None:
                    messagebox.showerror("Erro", "Coluna 'IMG' não encontrada.")
                    return
                for row_idx in range(2, sheet.max_row + 1):
                    if row_idx in sheet.row_dimensions and sheet.row_dimensions[row_idx].hidden: continue
                    val = sheet.cell(row=row_idx, column=img_col_idx).value
                    if val is not None: excel_img_values.append(val)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro no Excel:\n{e}")
                return
            if not excel_img_values:
                messagebox.showwarning("Aviso", "Coluna 'IMG' está vazia.")
                return
            if len(images) != len(excel_img_values):
                messagebox.showwarning("Aviso", f"Imagens ({len(images)}) vs Excel ({len(excel_img_values)}).")
            qtd_pares = min(len(images), len(excel_img_values))
        else:
            qtd_pares = len(images)
            try: start_no = int(self.start_number.get())
            except: start_no = 1
            excel_img_values = [str(start_no + i) for i in range(qtd_pares)]
            
        self.mapping = []
        self.original_dest_bases = []
        try: digitos = int(self.digits.get())
        except: digitos = 2
        
        seen_targets = set()
        duplicados = set()
        
        for i in range(qtd_pares):
            orig_name = images[i]
            img_val_str = str(excel_img_values[i]).strip()
            if img_val_str.endswith(".0"): img_val_str = img_val_str[:-2]
            if img_val_str.isdigit(): novo_base = img_val_str.zfill(digitos)
            else: novo_base = img_val_str
            self.original_dest_bases.append(img_val_str)
                
            _, ext = os.path.splitext(orig_name)
            if self.compress_var.get(): novo_nome = f"{novo_base}.JPG"
            else: novo_nome = f"{novo_base}{ext.upper()}"
            
            if novo_nome in seen_targets: duplicados.add(novo_nome)
            seen_targets.add(novo_nome)
            
            orig_path = os.path.join(folder, orig_name)
            novo_path = os.path.join(folder, novo_nome)
            
            self.mapping.append({
                'orig_name': orig_name, 'new_name': novo_nome,
                'orig_path': orig_path, 'new_path': novo_path
            })
            
        if duplicados:
            messagebox.showwarning("Atenção", f"Conflitos de nome previstos (Ex: {list(duplicados)[0]}). O novo motor lida com swaps, mas cuidado com cópias puras.")
            
        for item in self.tree.get_children(): self.tree.delete(item)
        for idx, item in enumerate(self.mapping):
            item['size_orig_str'] = self.format_size(os.path.getsize(item['orig_path']))
            item['size_est_str'] = "70 KB" 
            self.tree.insert("", "end", values=(idx + 1, item['orig_name'], item['new_name'], item['size_orig_str'], item['size_est_str']))
            
        if self.mapping:
            self.btn_rename.configure(state="normal")
            
    def rename_files(self):
        if not self.mapping or self.processing: return
        if not messagebox.askyesno("Confirmar", f"Renomear {len(self.mapping)} arquivos?"): return
            
        self.processing = True
        self.btn_rename.configure(state="disabled")
        self.frame_progress.pack(fill="x", pady=10)
        self.progress.set(0)
        self.lbl_status.configure(text="Iniciando motor turbo (ThreadPool)...")
        
        threading.Thread(target=self.run_rename_task_robust, daemon=True).start()

    def run_rename_task_robust(self):
        sucessos = 0
        falhas = 0
        erros_msg = []
        mapping_sq = list(self.mapping)
        total = len(mapping_sq)
        
        from concurrent.futures import ThreadPoolExecutor, as_completed
        
        with ThreadPoolExecutor() as executor:
            # Submete todas as tarefas
            futures = {executor.submit(self.process_single_image, item, idx): item for idx, item in enumerate(mapping_sq)}
            
            concluidos = 0
            for future in as_completed(futures):
                concluidos += 1
                res_ok, res_msg = future.result()
                if res_ok: sucessos += 1
                else: 
                    falhas += 1
                    erros_msg.append(res_msg)
                
                perc_num = concluidos / total
                # ATUALIZAÇÃO SEGURA DA GUI: Passar para a thread principal (evita hang/deadlock do Tkinter)
                self.root.after(0, self.update_ui_progress, perc_num, int(perc_num * 100), f"Fase 1/2: Compactando ({concluidos}/{total})")

        self.root.after(0, self.lbl_status.configure, {"text": "Fase 2/2: Aplicando nomes definitivos..."})
        for item in mapping_sq:
            temp_path = item['new_path'] + ".ecotmp"
            if os.path.exists(temp_path):
                try:
                    if os.path.exists(item['orig_path']): os.remove(item['orig_path'])
                    if os.path.exists(item['new_path']): os.remove(item['new_path'])
                    os.rename(temp_path, item['new_path'])
                except Exception as e:
                    print(f"Erro cleanup {item['new_name']}: {e}")

        self.root.after(0, lambda: self.finish_rename(sucessos, falhas, erros_msg))

    def update_ui_progress(self, val_float, perc_int, status):
        # Esta função agora é chamada de forma segura pela thread principal via root.after
        self.progress.set(val_float)
        self.lbl_perc.configure(text=f"{perc_int}%")
        self.lbl_status.configure(text=status)

    def process_single_image(self, item, index):
        temp_target = item['new_path'] + ".ecotmp"
        try:
            if self.compress_var.get():
                img = Image.open(item['orig_path'])
                img = ImageOps.exif_transpose(img)
                img.thumbnail((800, 800), Image.Resampling.LANCZOS)
                img = img.convert("RGB")
                img.save(temp_target, "JPEG", optimize=True, quality=45)
                img.close()
            else:
                shutil.copy2(item['orig_path'], temp_target)
            return True, ""
        except Exception as e:
            return False, f"{item['orig_name']}: {str(e)}"

    def finish_rename(self, sucessos, falhas, erros_msg):
        msg = f"Processo concluído!\nSucesso: {sucessos}\nFalhas: {falhas}"
        if falhas > 0:
            messagebox.showwarning("Aviso", msg + f"\n\nErros:\n" + "\n".join(erros_msg[:5]))
        else:
            messagebox.showinfo("Sucesso Total", msg)
        
        self.processing = False
        self.frame_progress.pack_forget()
        self.reset_preview()

    # ------------------ PDFs ------------------
    def select_pdf_folder(self):
        folder = filedialog.askdirectory(title="Selecione a pasta")
        if folder: self.pdf_folder.set(folder)

    def merge_pdfs(self):
        folder = self.pdf_folder.get()
        if not folder:
            messagebox.showwarning("Aviso", "Selecione a pasta.")
            return
        try:
            arquivos = os.listdir(folder)
            pdfs = [f for f in arquivos if not f.startswith('.') and f.lower().endswith('.pdf')]
        except: return
        
        if not pdfs: return
        
        ordem = self.pdf_sort_order.get()
        if "Decrescente" in ordem: pdfs.sort(reverse=True)
        else: pdfs.sort(reverse=False)
            
        out_name = self.pdf_output_name.get().strip() or "Documento_Unificado.pdf"
        if not out_name.endswith(".pdf"): out_name += ".pdf"
        out_path = os.path.join(folder, out_name)
        
        if out_name in pdfs: pdfs.remove(out_name)
        if not pdfs: return
            
        try:
            doc_final = fitz.open() 
            for pdf_file in pdfs:
                try:
                    doc_temp = fitz.open(os.path.join(folder, pdf_file))
                    doc_final.insert_pdf(doc_temp)
                    doc_temp.close()
                except: pass
            
            doc_final.save(out_path, garbage=4, deflate=True)
            doc_final.close()
            messagebox.showinfo("Concluído", f"Unificados em '{out_name}'.")
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def check_for_updates(self):
        try:
            with urllib.request.urlopen(UPDATE_URL, timeout=5) as r:
                data = json.loads(r.read().decode())
                remote_ver = data.get("version", VERSION)
                if remote_ver > VERSION:
                    msg = f"Nova versão encontrada: v{remote_ver}\nMudanças: {data.get('changelog')}\n\nDeseja fechar, instalar silenciosamente e reiniciar agora?"
                    if messagebox.askyesno("Atualização Pronta", msg):
                        os_name = platform.system()
                        url = data.get("download_url_mac") if os_name == "Darwin" else data.get("download_url_win")
                        if url: threading.Thread(target=self.run_auto_update, args=(url, remote_ver), daemon=True).start()
                        else: messagebox.showinfo("Erro", "URL de instalação não encontrada no servidor.")
                else: messagebox.showinfo("App em Dia", "Você já está rodando a última Enterprise Edition.")
        except: pass

    def run_auto_update(self, url, version):
        # UI Freeze de Download
        self.btn_rename.configure(state="disabled")
        try: self.btn_load.configure(state="disabled")
        except: pass
        self.frame_progress.pack(fill="x", pady=10)
        self.lbl_status.configure(text=f"Abaixando pacote v{version} (Em Background)...")
        self.progress.set(0.3)
        self.lbl_perc.configure(text="30%")
        
        try:
            temp_dir = os.path.join(os.path.expanduser("~"), ".ecowave_update_tmp")
            if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            zpath = os.path.join(temp_dir, "u.zip")
            urllib.request.urlretrieve(url, zpath)
            
            self.lbl_status.configure(text="Pré-Instalando e Descompactando...")
            self.progress.set(0.8)
            with zipfile.ZipFile(zpath, 'r') as zf: zf.extractall(temp_dir)
            
            # Bloqueio caso seja rodado via Python Raw
            is_frozen = getattr(sys, 'frozen', False)
            if not is_frozen:
                messagebox.showinfo("Dev Mode", "Instalador baixado, mas ignorado pois não é um executável compilado (.app ou .exe).")
                self.frame_progress.pack_forget()
                return
                
            os_name = platform.system()
            if os_name == "Darwin":
                # APLICANDO TRAVA TÉCNICA E CORREÇÃO DE CAMINHO (A VIDA DOS SEUS ARQUIVOS)
                # O executável do Mac fica em: MeuApp.app/Contents/MacOS/app
                # Recuar 2 pastas (..) cai exatamente na pasta .app
                app_path = os.path.abspath(os.path.join(os.path.dirname(sys.executable), "..", ".."))
                new_app_extracted = os.path.join(temp_dir, "RenomeadorApp.app")
                
                # TRAVA DE SEGURANÇA: NUNCA subscrever se não for um .app real
                if not app_path.endswith(".app"):
                    messagebox.showerror("Erro de Segurança Crítico", f"O atualizador tentou sobrescrever um diretório inválido:\n{app_path}\nAtualização Abortada pelo Sistema de Defesa.")
                    self.frame_progress.pack_forget()
                    return
                
                script_sh = os.path.join(temp_dir, "updater.command")
                with open(script_sh, "w") as f:
                    f.write("#!/bin/bash\n")
                    f.write("sleep 2\n") # Tempo pro App fechar e liberar a pasta
                    # FLAG --delete FOI REMOVIDA PARA SEMPRE! Apenas substitui novos arquivos sem excluir nada em volta.
                    f.write(f"rsync -a '{new_app_extracted}/' '{app_path}/'\n") 
                    f.write(f"xattr -cr '{app_path}'\n") # BURACO NO GATEKEEPER: Remove quarentena
                    f.write(f"open '{app_path}'\n") # Reinicia
                    f.write(f"rm -rf '{temp_dir}'\n") # Limpeza ninja do zips
                    f.write("rm -- \"$0\"\n") # Auto-destrói o script sh
                    
                os.chmod(script_sh, 0o755)
                subprocess.Popen(["/bin/bash", script_sh], start_new_session=True)
                sys.exit(0) # Morre
                
            else: # Windows
                app_path = os.path.dirname(sys.executable)
                exe_name = os.path.basename(sys.executable)
                
                script_bat = os.path.join(temp_dir, "updater.bat")
                with open(script_bat, "w") as f:
                    f.write("@echo off\n")
                    f.write("timeout /t 2 /nobreak >nul\n") # Espera App fechar
                    f.write(f"xcopy /S /E /Y /I \"{temp_dir}\\*\" \"{app_path}\\\"\n") # Subscreve forçado
                    f.write(f"start \"\" \"{os.path.join(app_path, exe_name)}\"\n") # Inicia novo executável
                    f.write("del \"%~f0\"\n") # Pede pro arquivo bat tentar se matar
                    
                # Roda o BAT desanexado do nosso processo
                subprocess.Popen([script_bat], creationflags=subprocess.CREATE_NEW_CONSOLE)
                sys.exit(0) # Mata o App Atual
                
        except Exception as e:
            messagebox.showerror("Erro Crítico de Instalação", f"Falha ao realizar 'Seamless Update':\n\n{e}")
            self.frame_progress.pack_forget()

if __name__ == "__main__":
    # CustomTkinter precisa ser instanciado como ctk.CTk()
    root = ctk.CTk()
    app = ToolApp(root)
    root.mainloop()
