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
    "Direito Trabalhista - CLT": {
        "Agravo TST": 8,
        "Agravo de Instrumento em RR ou RO": 8,
        "Agravo de Petição": 8,
        "Contraminuta AI / Contrarrazões RR": 8,
        "Contrarrazões ao RR ou RO": 8,
        "Embargos TST": 8,
        "Embargos à Execução": 5,
        "Embargos de declaração": 5,
        "Impugnação à Sentença de Liquidação": 5,
        "Recurso de Revista": 8,
        "Recurso Ordinário Trabalhista": 8
    },
    "Direito Civil - CPC": {
        "Agravo de Instrumento": 15,
        "Apelação": 15,
        "Embargos de declaração": 5,
        "Recurso Especial": 15,
        "Recurso Extraordinário": 15,
        "Recurso Ordinário (Cível)": 15
    }
}

feriados_selecionados = []

def calcular_prazo_util(data_str, dias):
    try:
        data_inicial = datetime.strptime(data_str, "%d/%m")
        data_inicial = data_inicial.replace(year=datetime.now().year)
        df = pd.date_range(data_inicial + timedelta(days=1), periods=90, freq='B')

        df_filtrado = [d for d in df if d.strftime("%d/%m") not in feriados_selecionados]

        if len(df_filtrado) < dias:
            return "Prazo ultrapassa os dias úteis disponíveis"

        termo_final = df_filtrado[dias - 1]
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

def alternar_feriado():
    if checkbox_var.get():
        feriado_frame.pack(pady=5)
    else:
        feriado_frame.pack_forget()
        feriados_selecionados.clear()

def adicionar_feriado():
    inicio = feriado_inicio.get()
    fim = feriado_fim.get()
    try:
        data_inicio = datetime.strptime(inicio, "%d/%m")
        data_fim = datetime.strptime(fim, "%d/%m")
        ano_atual = datetime.now().year
        data_inicio = data_inicio.replace(year=ano_atual)
        data_fim = data_fim.replace(year=ano_atual)
        datas = pd.date_range(data_inicio, data_fim).to_pydatetime()
        for d in datas:
            feriados_selecionados.append(d.strftime("%d/%m"))
        feriado_inicio.delete(0, tk.END)
        feriado_fim.delete(0, tk.END)
        messagebox.showinfo("Feriado adicionado", f"Feriado(s) registrado(s): {inicio} até {fim}")
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM.")

# Interface principal
janela = tk.Tk()
janela.title("Gerador de Prazos Jurídicos Digitais")
janela.geometry("700x700")
janela.configure(bg="#0f172a")

frame = tk.Frame(janela, bg="#0f172a")
frame.pack(expand=True)

# Exibição da imagem no topo
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

# Variáveis para seleção de ramo e tipo
ramo_var = tk.StringVar()
tipo_var = tk.StringVar()

# Campo de data da publicação
tk.Label(
    frame,
    text="Data da publicação (formato: DD/MM):",
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
tk.Label(frame, text="Selecione o ramo:", font=("Segoe UI", 11), bg="#0f172a", fg="white").pack(pady=(10, 0))
ramo_menu = ttk.Combobox(frame, textvariable=ramo_var, values=list(PRAZOS.keys()), state="readonly", width=40)
ramo_menu.pack(pady=2)

def atualizar_tipos(event):
    ramo = ramo_var.get()
    tipo_menu["values"] = sorted(PRAZOS.get(ramo, {}).keys())
    tipo_var.set("")

ramo_menu.bind("<<ComboboxSelected>>", atualizar_tipos)

tk.Label(frame, text="Selecione o tipo de recurso:", font=("Segoe UI", 11), bg="#0f172a", fg="white").pack(pady=(10, 0))
tipo_menu = ttk.Combobox(frame, textvariable=tipo_var, state="readonly", width=40)
tipo_menu.pack(pady=2)

# Botão de cálculo do prazo jurídico
tk.Button(
    frame,
    text="Calcular Prazo Jurídico",
    font=("Segoe UI", 11, "bold"),
    bg="#2563eb",
    fg="white",
    activebackground="#1d4ed8",
    command=exibir_calculo
).pack(pady=10)

# Interface visual de feriados
checkbox_var = tk.BooleanVar()
tk.Checkbutton(
    frame,
    text="Ao longo do prazo existe feriados?",
    variable=checkbox_var,
    command=alternar_feriado,
    bg="#0f172a",
    fg="white",
    activebackground="#0f172a",
    selectcolor="#0f172a",
    font=("Segoe UI", 11)
).pack(pady=(5, 0))

feriado_frame = tk.Frame(frame, bg="#0f172a")

tk.Label(
    feriado_frame,
    text="De (DD/MM):",
    font=("Segoe UI", 10),
    bg="#0f172a",
    fg="#94a3b8"
).grid(row=0, column=0, padx=5)
feriado_inicio = tk.Entry(feriado_frame, font=("Consolas", 12), width=10, justify="center", bg="#1e293b", fg="#f8fafc")
feriado_inicio.grid(row=0, column=1, padx=5)

tk.Label(
    feriado_frame,
    text="Até (DD/MM):",
    font=("Segoe UI", 10),
    bg="#0f172a",
    fg="#94a3b8"
).grid(row=0, column=2, padx=5)
feriado_fim = tk.Entry(feriado_frame, font=("Consolas", 12), width=10, justify="center", bg="#1e293b", fg="#f8fafc")
feriado_fim.grid(row=0, column=3, padx=5)

adicionar_feriado_btn = tk.Button(
    feriado_frame,
    text="Adicionar Feriado",
    font=("Segoe UI", 10, "bold"),
    bg="#334155",
    fg="white",
    command=adicionar_feriado
)
adicionar_feriado_btn.grid(row=0, column=4, padx=5)

tk.Label(
    frame,
    text="Aplicação desenvolvida por: Rodrigo Junqueira de Lima Siqueira",
    font=("Segoe UI", 8),
    bg="#0f172a",
    fg="#64748b"
).pack(side="bottom", pady=(10, 5))

janela.mainloop()








