import customtkinter as ctk
from tkinter import filedialog
from PIL import Image
import os
import sys
import subprocess

from ingestion.leitor_excel import ler_excel
from validation.validador import validar_dados
from core.calculadora import gerar_mapa_detalhado
from core.exportador import exportar_excel


# =========================
# CONFIG VISUAL
# =========================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


# =========================
# PATH BASE
# =========================
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.abspath(".")


BASE_DIR = get_base_path()

caminho_logo = os.path.join(BASE_DIR, "assets", "logo.png")
caminho_icon = os.path.join(BASE_DIR, "assets", "logo.ico")


# =========================
# FUNÇÕES
# =========================

def selecionar_arquivo():
    caminho = filedialog.askopenfilename(
        title="Selecionar ficheiro Excel",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if caminho:
        entry_arquivo.delete(0, "end")
        entry_arquivo.insert(0, caminho)
        atualizar_status("Arquivo selecionado", "lightblue")
        log(f"📂 Arquivo selecionado:\n{caminho}")


def log(msg):
    log_output.insert("end", msg + "\n")
    log_output.see("end")


def limpar_log():
    log_output.delete("1.0", "end")


def atualizar_status(msg, cor="white"):
    label_status.configure(text=f"● {msg}", text_color=cor)


def processar():
    caminho = entry_arquivo.get()

    if not caminho or not os.path.exists(caminho):
        atualizar_status("Arquivo inválido", "red")
        log("❌ Selecione um arquivo válido.")
        return

    try:
        limpar_log()
        atualizar_status("Processando...", "orange")
        botao_processar.configure(state="disabled")
        entry_arquivo.configure(state="disabled")

        log("📥 Lendo arquivo...")
        dados, ano, mes = ler_excel(caminho)

        log("🔍 Validando dados...")
        erros = validar_dados(dados, ano, mes)

        if erros:
            atualizar_status("Erros encontrados", "red")
            log("❌ ERROS DE VALIDAÇÃO:")
            for e in erros:
                log(f"Dia {e['dia']} | {e['colaborador']} → {e['erro']}")
            return

        log("🧮 Calculando mapa...")
        mapa = gerar_mapa_detalhado(dados)

        log("📊 Exportando Excel...")
        caminhos = exportar_excel(mapa, ano, mes)

        atualizar_status("Concluído com sucesso", "green")
        log("✅ Processo finalizado!")
        log("📁 Arquivos gerados:")

        for c in caminhos:
            log(c)

        if caminhos:
            pasta = os.path.dirname(caminhos[0])
            subprocess.Popen(f'explorer "{pasta}"')

    except Exception as e:
        atualizar_status("Erro inesperado", "red")
        log("❌ Erro inesperado:")
        log(str(e))

    finally:
        botao_processar.configure(state="normal")
        entry_arquivo.configure(state="normal")


# =========================
# APP
# =========================

app = ctk.CTk()
app.title("Unicâmbio - Sistema de Mapa de Horas")

# 🔥 TAMANHO FINAL AJUSTADO
app.geometry("750x650")
app.minsize(750, 600)

app.iconbitmap(caminho_icon)


# =========================
# GRID PRINCIPAL
# =========================
app.grid_rowconfigure(0, weight=1)
app.grid_rowconfigure(1, weight=0)
app.grid_columnconfigure(0, weight=1)

main_frame = ctk.CTkFrame(app)
main_frame.grid(row=0, column=0, sticky="nsew")


# =========================
# HEADER
# =========================

header = ctk.CTkFrame(main_frame, height=70, fg_color="#000000")
header.pack(fill="x")

logo_img = ctk.CTkImage(
    light_image=Image.open(caminho_logo),
    dark_image=Image.open(caminho_logo),
    size=(170, 45)
)

logo_label = ctk.CTkLabel(header, image=logo_img, text="")
logo_label.pack(pady=10)

linha = ctk.CTkFrame(main_frame, height=3, fg_color="#FFD100")
linha.pack(fill="x")


# =========================
# TÍTULO
# =========================

titulo = ctk.CTkLabel(
    main_frame,
    text="Gerador de Mapa de Horas",
    font=("Arial", 20, "bold")
)
titulo.pack(pady=15)


# =========================
# INPUT
# =========================

frame_arquivo = ctk.CTkFrame(main_frame)
frame_arquivo.pack(padx=20, pady=10, fill="x")

entry_arquivo = ctk.CTkEntry(
    frame_arquivo,
    placeholder_text="Selecione a escala mensal (.xlsx)"
)
entry_arquivo.pack(side="left", fill="x", expand=True, padx=10, pady=10)

btn_procurar = ctk.CTkButton(
    frame_arquivo,
    text="Procurar",
    command=selecionar_arquivo,
    width=120
)
btn_procurar.pack(side="right", padx=10)


# =========================
# STATUS
# =========================

label_status = ctk.CTkLabel(
    main_frame,
    text="● Aguardando arquivo",
    font=("Arial", 14, "bold")
)
label_status.pack(pady=10)


# =========================
# LOG
# =========================

frame_log = ctk.CTkFrame(main_frame)
frame_log.pack(padx=20, pady=10, fill="both", expand=True)

log_output = ctk.CTkTextbox(frame_log)
log_output.pack(fill="both", expand=True, padx=10, pady=10)


# =========================
# BOTÃO
# =========================

botao_processar = ctk.CTkButton(
    main_frame,
    text="Gerar Mapas",
    command=processar,
    height=50,
    font=("Arial", 15, "bold"),
    fg_color="#FFD100",
    text_color="black",
    hover_color="#e6bc00"
)
botao_processar.pack(pady=10)


# =========================
# RODAPÉ FIXO
# =========================

rodape = ctk.CTkLabel(
    app,
    text="© 2026 Daniel Viana • Sistema de Mapa de Horas",
    font=("Arial", 10),
    text_color="gray"
)
rodape.grid(row=1, column=0, pady=5)


# =========================
# START
# =========================

app.mainloop()