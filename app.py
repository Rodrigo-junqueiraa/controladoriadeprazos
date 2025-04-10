import os
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from openpyxl import load_workbook
from PIL import Image, ImageTk
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

print("Iniciando app...")

ARQUIVO_MODELO = "Planilha de prazos - atualizada.xlsx"

# Função para garantir caminho correto da imagem no .exe
def recurso_path(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)

IMAGEM_JURIDICA = recurso_path("justica.png")

# Prazos por tipo e ramo
PRAZOS = {
    "CLT": {
        "Embargos de declaração": 5,
        "Recurso Ordinário Trabalhista": 8,
        "Recurso de Revista": 8,
        "Agravo de Petição": 8,
        "Agravo de Instrumento em RR ou RO": 8,
        "Contrarrazões ao RR ou RO": 8,
        "Contraminuta ao RR / Contrarrazões ao AI": 8,
    },
    "CPC": {
        "Embargos de declaração": 5,
        "Recurso Ordinário (Cível)": 15,
        "Apelação": 15,
        "Agravo de Instrumento": 15,
        "Recurso Especial": 15,
        "Recurso Extraordinário": 15,
    }
}

def calcular_prazo_util(data_str, dias):
    try:
        data_inicial = datetime.strptime(data_str, "%d/%m")
        data_inicial = data_inicial.replace(year=datetime.now().year)
        df = pd.date_range(data_inicial + timedelta(days=1), periods=60, freq='B')
        termo_final = df[dias - 1]
        return termo_final.strftime("%d/%m")
    except Exception as e:
        return f"Erro: {e}"

def exibir_calculo():
    data = publicacao_entry.get()
    ramo = ramo_var.get()
    tipo = tipo_var.get()
    if ramo in PRAZOS and tipo in PRAZOS[ramo]:
        prazo = PRAZOS[ramo][tipo]
        resultado = calcular_prazo_util(data, prazo)
        messagebox.showinfo("Resultado do Cálculo", f"{tipo}\nPrazo final: {resultado}")
    else:
        messagebox.showerror("Erro", "Selecione um tipo e ramo válido.")

def preencher_pz(data, janela):
    if not os.path.exists(ARQUIVO_MODELO):
        messagebox.showerror("Erro", f"Arquivo '{ARQUIVO_MODELO}' não encontrado.", parent=janela)
        return

    try:
        wb = load_workbook(ARQUIVO_MODELO)
        ws = wb.active

        for row in ws.iter_rows(min_row=1):
            if row[0].value == "TIPO":
                row[3].value = data

        nome_arquivo = f"Planilha_prazos_{data.replace('/', '-')}.xlsx"
        salvar_em = filedialog.asksaveasfilename(
            parent=janela,
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=nome_arquivo,
            title="Salvar como"
        )

        if salvar_em:
            wb.save(salvar_em)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{salvar_em}", parent=janela)

    except Exception as e:
        messagebox.showerror("Erro ao processar", str(e), parent=janela)

# Interface
janela = tk.Tk()
janela.title("Gerador de Prazos Jurídicos Digitais")
janela.geometry("600x650")
janela.configure(bg="#0f172a")

frame = tk.Frame(janela, bg="#0f172a")
frame.pack(expand=True)

# Imagem no topo
if os.path.exists(IMAGEM_JURIDICA):
    try:
        imagem = Image.open(IMAGEM_JURIDICA)
        imagem = imagem.resize((100, 100), Image.LANCZOS)
        imagem_tk = ImageTk.PhotoImage(imagem)
        img_label = tk.Label(frame, image=imagem_tk, bg="#0f172a")
        img_label.image = imagem_tk
        img_label.pack(pady=(15, 5))
    except Exception as e:
        print("Erro ao carregar imagem:", e)

# Título
tk.Label(
    frame,
    text="Sistema de Controle de Prazos",
    font=("Segoe UI", 18, "bold"),
    bg="#0f172a",
    fg="#38bdf8"
).pack(pady=(5, 10))

# Instrução data para preenchimento
entrada_label = tk.Label(
    frame,
    text="Insira a data do prazo (formato: DD/MM):",
    font=("Segoe UI", 12),
    bg="#0f172a",
    fg="#94a3b8"
)
entrada_label.pack()

data_label = tk.Entry(
    frame,
    font=("Consolas", 14),
    width=20,
    justify="center",
    bd=2,
    relief="flat",
    bg="#1e293b",
    fg="#f8fafc",
    insertbackground="#f8fafc"
)
data_label.pack(pady=8)
entrada = data_label

# Botão gerar planilha
tk.Button(
    frame,
    text="Gerar Arquivo Excel",
    font=("Segoe UI", 12, "bold"),
    bg="#2563eb",
    fg="white",
    activebackground="#1d4ed8",
    command=lambda: preencher_pz(entrada.get(), janela)
).pack(pady=10)

# Instrução para cálculo jurídico
tk.Label(
    frame,
    text="Insira a data da publicação (formato: DD/MM):",
    font=("Segoe UI", 12),
    bg="#0f172a",
    fg="#94a3b8"
).pack()

publicacao_entry = tk.Entry(
    frame,
    font=("Consolas", 14),
    width=20,
    justify="center",
    bd=2,
    relief="flat",
    bg="#1e293b",
    fg="#f8fafc",
    insertbackground="#f8fafc"
)
publicacao_entry.pack(pady=8)

# Seleção de ramo e tipo
ramo_var = tk.StringVar()
tipo_var = tk.StringVar()

tk.Label(frame, text="Selecione o ramo:", font=("Segoe UI", 11), bg="#0f172a", fg="white").pack(pady=(10, 0))
ramo_menu = ttk.Combobox(frame, textvariable=ramo_var, values=list(PRAZOS.keys()), state="readonly")
ramo_menu.pack(pady=2)

def atualizar_tipos(event):
    ramo = ramo_var.get()
    tipo_menu["values"] = list(PRAZOS.get(ramo, {}).keys())
    tipo_var.set("")

ramo_menu.bind("<<ComboboxSelected>>", atualizar_tipos)

tk.Label(frame, text="Selecione o tipo de recurso:", font=("Segoe UI", 11), bg="#0f172a", fg="white").pack(pady=(10, 0))
tipo_menu = ttk.Combobox(frame, textvariable=tipo_var, state="readonly")
tipo_menu.pack(pady=2)

# Botão calcular prazo
tk.Button(
    frame,
    text="Calcular Prazo Jurídico",
    font=("Segoe UI", 11, "bold"),
    bg="#2563eb",
    fg="white",
    activebackground="#1d4ed8",
    command=exibir_calculo
).pack(pady=10)

# Assinatura
tk.Label(
    frame,
    text="Programa desenvolvido por: Rodrigo Junqueira de Lima Siqueira",
    font=("Segoe UI", 8),
    bg="#0f172a",
    fg="#64748b"
).pack(side="bottom", pady=(10, 5))

print("Interface carregada.")
janela.mainloop()
