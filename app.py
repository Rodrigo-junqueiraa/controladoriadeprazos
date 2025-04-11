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

def recurso_path(rel_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)

IMAGEM_JURIDICA = recurso_path("justica.png")

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
registros = []

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

def abrir_janela_preenchimento_multiplo():
    nova_janela = tk.Toplevel(janela)
    nova_janela.title("Múltiplos Preenchimentos")
    nova_janela.geometry("500x600")
    nova_janela.configure(bg="#0f172a")

    tk.Label(nova_janela, text="Data do Prazo (DD/MM):", bg="#0f172a", fg="white").pack()
    entry_data_prazo = tk.Entry(nova_janela)
    entry_data_prazo.pack()

    tk.Label(nova_janela, text="Nome:", bg="#0f172a", fg="white").pack()
    entry_nome = tk.Entry(nova_janela)
    entry_nome.pack()

    tk.Label(nova_janela, text="Processo:", bg="#0f172a", fg="white").pack()
    entry_processo = tk.Entry(nova_janela)
    entry_processo.pack()

    tk.Label(nova_janela, text="Tipo de Prazo:", bg="#0f172a", fg="white").pack()
    entry_tipo = tk.Entry(nova_janela)
    entry_tipo.pack()

    tk.Label(nova_janela, text="Responsável:", bg="#0f172a", fg="white").pack()
    entry_resp = tk.Entry(nova_janela)
    entry_resp.pack()

    tk.Label(nova_janela, text="Publicação:", bg="#0f172a", fg="white").pack()
    entry_pub = tk.Entry(nova_janela)
    entry_pub.pack()

    tree = ttk.Treeview(nova_janela, columns=("Nome", "Processo", "Tipo", "Data", "Resp", "Pub"), show="headings")
    for col in ("Nome", "Processo", "Tipo", "Data", "Resp", "Pub"):
        tree.heading(col, text=col)
    tree.pack(pady=5, fill="both", expand=True)

    def adicionar():
        nome = entry_nome.get()
        processo = entry_processo.get()
        tipo = entry_tipo.get()
        data = entry_data_prazo.get()
        resp = entry_resp.get()
        pub = entry_pub.get()
        if nome and processo and tipo and data and resp and pub:
            registros.append((nome, processo, tipo, data, resp, pub))
            tree.insert("", "end", values=(nome, processo, tipo, data, resp, pub))
            entry_nome.delete(0, tk.END)
            entry_processo.delete(0, tk.END)
            entry_tipo.delete(0, tk.END)
            entry_resp.delete(0, tk.END)
            entry_pub.delete(0, tk.END)

    def gerar_excel():
        if not registros:
            messagebox.showerror("Erro", "Nenhum dado para preencher.")
            return
        try:
            wb = load_workbook(ARQUIVO_MODELO)
            ws = wb.active
            bloco_index = 0
            for i, row in enumerate(ws.iter_rows(min_row=1)):
                if bloco_index >= len(registros):
                    break
                if row[0].value and str(row[0].value).strip().upper() == "NOME" and row[1].value in (None, ""):
                    nome, processo, tipo, data, resp, pub = registros[bloco_index]
                    ws.cell(row=i+1, column=2).value = nome
                    ws.cell(row=i+2, column=2).value = processo
                    ws.cell(row=i+3, column=2).value = tipo
                    ws.cell(row=i+1, column=4).value = resp
                    ws.cell(row=i+2, column=4).value = pub
                    ws.cell(row=i+3, column=4).value = data
                    bloco_index += 1
            salvar_em = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if salvar_em:
                wb.save(salvar_em)
                messagebox.showinfo("Sucesso", f"Arquivo salvo em: {salvar_em}")
        except Exception as e:
            messagebox.showerror("Erro ao preencher planilha", str(e))

    tk.Button(nova_janela, text="Adicionar", bg="#2563eb", fg="white", command=adicionar).pack(pady=5)
    tk.Button(nova_janela, text="Preencher Dados Cliente/Processo", bg="#22c55e", fg="white", command=gerar_excel).pack(pady=10)
    tk.Label(nova_janela, text="Aplicação desenvolvida por: Rodrigo Junqueira de Lima Siqueira", font=("Segoe UI", 8), bg="#0f172a", fg="#64748b").pack(side="bottom", pady=(10, 5))

# Interface principal
janela = tk.Tk()
janela.title("Sistema de Controle de Prazos")
janela.geometry("700x700")
janela.configure(bg="#0f172a")

frame = tk.Frame(janela, bg="#0f172a")
frame.pack(expand=True)

# Imagem
if os.path.exists(IMAGEM_JURIDICA):
    try:
        imagem = Image.open(IMAGEM_JURIDICA)
        imagem = imagem.resize((100, 100), Image.LANCZOS)
        imagem_tk = ImageTk.PhotoImage(imagem)
        img_label = tk.Label(frame, image=imagem_tk, bg="#0f172a")
        img_label.image = imagem_tk
        img_label.pack(pady=(10, 5))
    except Exception as e:
        print("Erro ao carregar imagem:", e)

# Título principal
tk.Label(
    frame,
    text="Sistema de Controle de Prazos",
    font=("Segoe UI", 16, "bold"),
    bg="#0f172a",
    fg="white"
).pack(pady=(5, 5))

# Botão de preenchimento
btn_preencher = tk.Button(
    frame,
    text="Preencher Cliente / Processo",
    font=("Segoe UI", 12, "bold"),
    bg="#22c55e",
    fg="white",
    activebackground="#16a34a",
    command=abrir_janela_preenchimento_multiplo
)
btn_preencher.pack(pady=10)

# Separador de seção
tk.Label(
    frame,
    text="Calculadora de Prazos",
    font=("Segoe UI", 14, "bold"),
    bg="#0f172a",
    fg="white"
).pack(pady=(20, 5))

# Campo de data da publicação
ramo_var = tk.StringVar()
tipo_var = tk.StringVar()

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

ramo_menu = ttk.Combobox(frame, textvariable=ramo_var, values=list(PRAZOS.keys()), state="readonly", width=40)
ramo_menu.pack(pady=2)

def atualizar_tipos(event):
    ramo = ramo_var.get()
    tipo_menu["values"] = sorted(PRAZOS.get(ramo, {}).keys())
    tipo_var.set("")

ramo_menu.bind("<<ComboboxSelected>>", atualizar_tipos)



tipo_menu = ttk.Combobox(frame, textvariable=tipo_var, state="readonly", width=40)
tipo_menu.pack(pady=2)

# Botão para calcular prazo
calcular_btn = tk.Button(
    frame,
    text="Calcular Prazo Jurídico",
    font=("Segoe UI", 11, "bold"),
    bg="#2563eb",
    fg="white",
    activebackground="#1d4ed8",
    command=exibir_calculo
)
calcular_btn.pack(pady=10)

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

feriado_inicio = tk.Entry(feriado_frame, font=("Consolas", 12), width=10, justify="center", bg="#1e293b", fg="#f8fafc")
feriado_inicio.grid(row=0, column=1, padx=5)

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

