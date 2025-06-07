"""
VERSÃO 1.0.0

Este script converte um arquivo PDF para um arquivo DOCX usando Pandoc.
Requisitos:
- Python 3.x
- Biblioteca PyMuPDF (fitz)
- Biblioteca Pandoc (instalado e no PATH)

Sugestões de melhorias, favor entrar em contato:
https://github.com/augustodugusto
"""
import fitz  # PyMuPDF
import os
import subprocess
import webbrowser
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, font
import sys
import logging

# Configurar o logging no início do script
log_file = os.path.join(os.getenv('APPDATA'), 'ConversorPDF', 'conversor.log')
os.makedirs(os.path.dirname(log_file), exist_ok=True)
logging.basicConfig(filename=log_file, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Substitua todos os `print("[DEBUG]...")` por:
logging.info("Aplicação iniciada.")
# E em blocos de erro:
logging.error("Ocorreu uma falha na conversão.", exc_info=True)

def get_resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funcionando tanto em dev quanto no PyInstaller. """
    try:
       
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


PANDOC_PATH = get_resource_path(os.path.join("pandoc", "pandoc.exe"))
# -----------------------------

def converter_pdf_para_html_semantico(caminho_pdf: str, caminho_html: str, progress_var, root) -> bool:
    """Converte PDF para HTML semântico usando PyMuPDF."""
    try:
        doc_pdf = fitz.open(caminho_pdf)
        total_paginas = doc_pdf.page_count
        html_completo = ""
        for idx, pagina in enumerate(doc_pdf):
            html_completo += pagina.get_text("html")
            progresso_html = int(((idx + 1) / total_paginas) * 50)
            progress_var.set(progresso_html)
            root.after(0, root.update_idletasks)

        with open(caminho_html, "w", encoding="utf-8") as f:
            f.write(html_completo)
        doc_pdf.close()
        return True
    except Exception as e:
        messagebox.showerror("Erro ao gerar HTML", f"Ocorreu um erro na Fase 1 (PDF -> HTML):\n{e}")
        return False

def converter_html_para_docx_com_pandoc(caminho_html: str, caminho_docx: str, progress_var, root) -> bool:
    """Chama o Pandoc para converter o HTML em DOCX usando o caminho estático."""
    
    if not os.path.exists(PANDOC_PATH):
        messagebox.showerror(
            "Pandoc não encontrado",
            f"O executável do Pandoc não foi encontrado no caminho especificado:\n\n{PANDOC_PATH}\n\n"
            "Por favor, edite a variável PANDOC_PATH no topo do script."
        )
        return False

    cmd = [PANDOC_PATH, "-f", "html", "-t", "docx", caminho_html, "-o", caminho_docx]
    print("[DEBUG] Chamando Pandoc com comando:", " ".join(cmd))
    progress_var.set(60)
    root.after(0, root.update_idletasks)

    try:
        processo = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        if processo.returncode != 0:
            messagebox.showerror("Erro na conversão com Pandoc", f"Pandoc retornou um erro:\n\n{processo.stderr.decode('utf-8', 'ignore')}")
            return False
        if not os.path.exists(caminho_docx) or os.path.getsize(caminho_docx) == 0:
            messagebox.showerror("Documento inválido", "O Pandoc não gerou um arquivo .docx ou o arquivo está vazio.")
            return False
        progress_var.set(100)
        return True
    except Exception as e:
        messagebox.showerror("Erro inesperado no Pandoc", f"Ocorreu um erro na Fase 2 (HTML -> DOCX):\n{e}")
        return False

def run_conversion_process(progress_var, root, btn_selecionar, status_label):
    """Função principal da conversão, executada em uma thread."""
    caminho_html_temporario = None
    try:
        root.after(0, lambda: btn_selecionar.config(state='disabled', text="Convertendo..."))
        root.after(0, lambda: status_label.config(text="Status: Aguardando seleção de arquivo..."))

        caminho_pdf = filedialog.askopenfilename(
            title="Selecione o arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")]
        )
        if not caminho_pdf:
            return

        root.after(0, lambda: status_label.config(text="Status: Iniciando conversão..."))
        pasta = os.path.dirname(caminho_pdf)
        nome_base = os.path.splitext(os.path.basename(caminho_pdf))[0]
        caminho_docx = os.path.join(pasta, nome_base + " (Editável).docx")

        with tempfile.NamedTemporaryFile(suffix=".html", delete=False, mode='w', encoding='utf-8') as tmp_html:
            caminho_html_temporario = tmp_html.name
        
        progress_var.set(0)

        if not converter_pdf_para_html_semantico(caminho_pdf, caminho_html_temporario, progress_var, root):
            raise Exception("Falha na Fase 1: PDF para HTML")

      
        if not converter_html_para_docx_com_pandoc(caminho_html_temporario, caminho_docx, progress_var, root):
            raise Exception("Falha na Fase 2: HTML para DOCX")

        webbrowser.open(f"file://{os.path.realpath(caminho_docx)}")
        root.after(0, lambda: messagebox.showinfo("Sucesso", f"DOCX editável gerado em:\n{caminho_docx}"))
        root.after(0, lambda: status_label.config(text="Status: Conversão concluída com sucesso!"))

    except Exception as e:
        print(f"[DEBUG] Erro no processo: {e}")
        root.after(0, lambda: status_label.config(text=f"Status: Erro! Verifique o terminal para detalhes."))
    finally:
        if caminho_html_temporario and os.path.exists(caminho_html_temporario):
            os.remove(caminho_html_temporario)
        progress_var.set(0)
        root.after(0, lambda: btn_selecionar.config(state='normal', text="Selecionar PDF e Converter"))
        if "Erro" not in status_label.cget("text"):
             root.after(0, lambda: status_label.config(text="Status: Pronto."))

def start_conversion_thread(progress_var, root, btn, status_label):
    """Inicia a conversão em uma nova thread para não congelar a UI."""
    thread = threading.Thread(target=run_conversion_process, args=(progress_var, root, btn, status_label))
    thread.daemon = True
    thread.start()

# Interface simples com tkinter
root = tk.Tk()
root.title("Conversor PDF para DOCX")
root.geometry("550x320")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Fontes
font_titulo = font.Font(family="Segoe UI", size=16, weight="bold")
font_normal = font.Font(family="Segoe UI", size=10)
font_status = font.Font(family="Segoe UI", size=9, slant="italic")
font_botao = font.Font(family="Segoe UI", size=11, weight="bold")


frame = tk.Frame(root, padx=25, pady=20, bg="#f0f0f0")
frame.pack(expand=True, fill="both")

# Título
label_titulo = tk.Label(frame, text="Conversor PDF para DOCX", font=font_titulo, bg="#f0f0f0")
label_titulo.pack(pady=(0, 15))


label_instrucoes = tk.Label(
    frame,
    text="Converta seus arquivos PDF em documentos Word (.docx) editáveis\ncom um único clique.",
    font=font_normal,
    justify="center",
    bg="#f0f0f0"
)
label_instrucoes.pack(pady=(0, 20))

# Barra de progresso
progress_var = tk.IntVar(value=0)
progress_bar = ttk.Progressbar(
    frame,
    orient="horizontal",
    length=500,
    mode="determinate",
    maximum=100,
    variable=progress_var
)
progress_bar.pack(pady=10, fill='x')

# Frame para os botões (para alinhamento lado a lado)
frame_botoes = tk.Frame(frame, bg="#f0f0f0")
frame_botoes.pack(pady=(15, 20))

# Botão principal de conversão
btn_selecionar = tk.Button(
    frame_botoes,
    text="Selecionar PDF e Converter",
    font=font_botao,
    bg="#28a745",  # Verde sucesso
    fg="white",
    relief="flat",
    padx=20,
    pady=10,
    activebackground="#218838",
    activeforeground="white"
)
btn_selecionar.pack(side="left", padx=(0, 10))

# Botão de sair
btn_sair = tk.Button(
    frame_botoes,
    text="Sair",
    font=font_normal,
    command=root.destroy,
    relief="flat",
    bg="#f0f0f0",
    padx=20,
    pady=10
)
btn_sair.pack(side="left")

# Label de status na parte inferior
status_label = tk.Label(
    root,
    text="Status: Pronto.",
    font=font_status,
    relief="sunken",
    bd=1,
    bg="#e9ecef",
    anchor="w",
    padx=10
)
status_label.pack(side="bottom", fill="x")

btn_selecionar['command'] = lambda: start_conversion_thread(progress_var, root, btn_selecionar, status_label)

print("[DEBUG] Aplicação iniciada com caminho estático para Pandoc.")
if not os.path.exists(PANDOC_PATH):
    print(f"[AVISO] Pandoc não encontrado em: {PANDOC_PATH}. O programa irá falhar na conversão.")
    status_label.config(text=f"Status: Erro! Pandoc não encontrado no caminho configurado.")
    messagebox.showwarning("Configuração Inválida", f"O caminho para o Pandoc parece estar incorreto. Verifique a variável PANDOC_PATH no código.\n\nCaminho configurado:\n{PANDOC_PATH}")


root.mainloop()