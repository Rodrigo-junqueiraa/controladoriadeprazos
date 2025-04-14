import os
import sys
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from openpyxl import load_workbook
from PIL import Image, ImageTk
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import json
from datetime import datetime

import sys

def caminho_json():
    if getattr(sys, 'frozen', False):
        # Executável (.exe)
        base_path = os.path.dirname(sys.executable)
    else:
        # Modo desenvolvimento (.py)
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, "CAMINHO_JSON")

CAMINHO_JSON = caminho_json()


# Função para carregar os prazos salvos
def carregar_prazos():
    if not os.path.exists(CAMINHO_JSON):
        return []
    with open(CAMINHO_JSON, "r", encoding="utf-8") as f:
        return json.load(f)

# Função para salvar os prazos no JSON
def salvar_prazos(prazos):
    with open(CAMINHO_JSON, "w", encoding="utf-8") as f:
        json.dump(prazos, f, indent=4, ensure_ascii=False)

# Função para registrar novos prazos no JSON
# Ela será chamada dentro do gerar_excel()
def registrar_prazos_em_json():
    prazos_existentes = carregar_prazos()
    novos_registros = []
    for nome, processo, tipo, data, resp, pub, obs in registros:
        novo = {
            "cliente": nome,
            "processo": processo,
            "tipo_prazo": tipo,
            "data_fatal": data,
            "notificado": False,
            "registro_em": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
        novos_registros.append(novo)

    prazos_existentes.extend(novos_registros)
    salvar_prazos(prazos_existentes)

# Verifica se há prazos para hoje e avisa
# Pode ser chamado ao abrir o sistema ou com um botão próprio
def verificar_prazos_hoje():
    hoje = datetime.now().strftime("%d/%m")
    prazos = carregar_prazos()
    alertas = []
    for prazo in prazos:
        if not prazo["notificado"] and prazo["data_fatal"] == hoje:
            alertas.append(prazo)
            prazo["notificado"] = True

    if alertas:
        texto = "\n".join([f"{p['cliente']} - {p['tipo_prazo']} - {p['processo']}" for p in alertas])
        messagebox.showinfo("⚠️ Prazos para hoje", f"Os seguintes prazos vencem hoje:\n\n{texto}")

    salvar_prazos(prazos)

# Chamada da verificação automática ao abrir o app
verificar_prazos_hoje()


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
    nova_janela.geometry("520x650")
    nova_janela.configure(bg="#0f172a")

    def criar_label(texto):
        return tk.Label(nova_janela, text=texto, bg="#0f172a", fg="white")

    criar_label("Data do Prazo (DD/MM):").pack()
    entry_data_prazo = tk.Entry(nova_janela)
    entry_data_prazo.pack()

    criar_label("Nome:").pack()
    entry_nome = tk.Entry(nova_janela)
    entry_nome.pack()

    criar_label("Processo:").pack()
    entry_processo = tk.Entry(nova_janela)
    entry_processo.pack()

    criar_label("Tipo de Prazo:").pack()
    entry_tipo = tk.Entry(nova_janela)
    entry_tipo.pack()

    criar_label("Responsável:").pack()
    entry_resp = tk.Entry(nova_janela)
    entry_resp.pack()

    criar_label("Publicação:").pack()
    entry_pub = tk.Entry(nova_janela)
    entry_pub.pack()

    criar_label("Observações:").pack()
    entry_obs = tk.Text(nova_janela, height=3, wrap="word")
    entry_obs.pack(fill="x", padx=10, pady=5)

    colunas = ("Nome", "Processo", "Tipo", "Data", "Resp", "Pub", "Obs")
    tree = ttk.Treeview(nova_janela, columns=colunas, show="headings")
    for col in colunas:
        tree.heading(col, text=col)
    tree.pack(pady=5, fill="both", expand=True)

    def adicionar():
        nome = entry_nome.get()
        processo = entry_processo.get()
        tipo = entry_tipo.get()
        data = entry_data_prazo.get()
        resp = entry_resp.get()
        pub = entry_pub.get()
        obs = entry_obs.get("1.0", tk.END).strip()
        if all([nome, processo, tipo, data, resp, pub]):
            registros.append((nome, processo, tipo, data, resp, pub, obs))
            tree.insert("", "end", values=(nome, processo, tipo, data, resp, pub, obs[:30] + ("..." if len(obs) > 30 else "")))
            entry_nome.delete(0, tk.END)
            entry_processo.delete(0, tk.END)
            entry_tipo.delete(0, tk.END)
            entry_resp.delete(0, tk.END)
            entry_pub.delete(0, tk.END)
            entry_obs.delete("1.0", tk.END)

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
                    nome, processo, tipo, data, resp, pub, obs = registros[bloco_index]
                    ws.cell(row=i+1, column=2).value = nome
                    ws.cell(row=i+2, column=2).value = processo
                    ws.cell(row=i+3, column=2).value = tipo
                    ws.cell(row=i+1, column=4).value = resp
                    ws.cell(row=i+2, column=4).value = pub
                    ws.cell(row=i+3, column=4).value = data

                    # Quebra observações em 3 partes para preencher F, G, H
                    obs_linhas = obs.strip().splitlines()
                    texto_total = " ".join(obs_linhas)
                    partes = [texto_total[i:i+50] for i in range(0, len(texto_total), 50)]
                    for j in range(3):
                        if j < len(partes):
                            ws.cell(row=i+1+j, column=6).value = partes[j]

                    bloco_index += 1

            salvar_em = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if salvar_em:
              wb.save(salvar_em)

            # Salvar prazos no JSON
            registrar_prazos_em_json()

            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {salvar_em}")

        except Exception as e:
            messagebox.showerror("Erro ao preencher planilha", str(e))

    tk.Button(nova_janela, text="Adicionar", bg="#2563eb", fg="white", command=adicionar).pack(pady=5)
    tk.Button(nova_janela, text="Preencher Dados Cliente/Processo", bg="#22c55e", fg="white", command=gerar_excel).pack(pady=10)
    criar_label("Aplicação desenvolvida por: Rodrigo Junqueira de Lima Siqueira").pack(side="bottom", pady=(10, 5))

# Interface principal com abas no topo
janela = tk.Tk()
janela.title("Sistema de Controle de Prazos")
janela.geometry("850x650")
janela.configure(bg="#0f172a")

style = ttk.Style()
style.theme_use('default')
style.configure("TNotebook", background="#0f172a", borderwidth=0)
style.configure("TNotebook.Tab", background="#1e293b", foreground="white", font=("Segoe UI", 10, "bold"), padding=(20, 10))
style.map("TNotebook.Tab", background=[("selected", "#2563eb")], foreground=[("selected", "white")])

# Carregar ícones brancos
icon_inicio = tk.PhotoImage(file=recurso_path("icon_inicio.png")).subsample(40, 40)
icon_calc = tk.PhotoImage(file=recurso_path("icon_calc.png")).subsample(40, 40)
icon_notification = tk.PhotoImage(file=recurso_path("icon_notification.png")).subsample(40, 40)
icon_listagem = tk.PhotoImage(file=recurso_path("icon_listagem.png")).subsample(40, 40)
icon_config = tk.PhotoImage(file=recurso_path("icon_config.png")).subsample(40, 40)

abas = ttk.Notebook(janela, style="TNotebook")
abas.pack(fill="both", expand=True)

# === Aba Principal ===
aba_principal = tk.Frame(abas, bg="#0f172a")
abas.add(aba_principal, text="  Início  ", image=icon_inicio, compound="left")

# Imagem
if os.path.exists(IMAGEM_JURIDICA):
    try:
        imagem = Image.open(IMAGEM_JURIDICA)
        imagem = imagem.resize((140, 140), Image.LANCZOS)
        imagem_tk = ImageTk.PhotoImage(imagem)
        img_label = tk.Label(aba_principal, image=imagem_tk, bg="#0f172a")
        img_label.image = imagem_tk
        img_label.pack(pady=(80, 5), anchor="center")
    except Exception as e:
        print("Erro ao carregar imagem:", e)

# Título
tk.Label(aba_principal, text="Sistema de Controle de Prazos", font=("Segoe UI", 16, "bold"), bg="#0f172a", fg="white").pack(pady=(10, 5), anchor="center")

# Botão preencher
tk.Button(aba_principal, text="Preencher Cliente / Processo", font=("Segoe UI", 12, "bold"), bg="#22c55e", fg="white", command=abrir_janela_preenchimento_multiplo).pack(pady=(10, 30), anchor="center")

# === Aba Calculadora ===
aba_calculadora = tk.Frame(abas, bg="#0f172a")
abas.add(aba_calculadora, text="  Calculadora de Prazos  ", image=icon_calc, compound="left")

# Campos da calculadora
publicacao_entry = tk.Entry(aba_calculadora, font=("Consolas", 14), width=20, justify="center", bd=2, relief="flat", bg="#1e293b", fg="#f8fafc", insertbackground="#f8fafc")
publicacao_entry.pack(pady=8)

ramo_var = tk.StringVar()
tipo_var = tk.StringVar()
ramo_menu = ttk.Combobox(aba_calculadora, textvariable=ramo_var, values=list(PRAZOS.keys()), state="readonly", width=40)
ramo_menu.pack(pady=2)

def atualizar_tipos(event):
    ramo = ramo_var.get()
    tipo_menu["values"] = sorted(PRAZOS.get(ramo, {}).keys())
    tipo_var.set("")

ramo_menu.bind("<<ComboboxSelected>>", atualizar_tipos)

tipo_menu = ttk.Combobox(aba_calculadora, textvariable=tipo_var, state="readonly", width=40)
tipo_menu.pack(pady=2)

# Botão calcular
tk.Button(aba_calculadora, text="Calcular Prazo Jurídico", font=("Segoe UI", 11, "bold"), bg="#2563eb", fg="white", command=exibir_calculo).pack(pady=10)

# Checkbox feriado
checkbox_var = tk.BooleanVar()
tk.Checkbutton(aba_calculadora, text="Ao longo do prazo existe feriados?", variable=checkbox_var, command=alternar_feriado, bg="#0f172a", fg="white", selectcolor="#0f172a", font=("Segoe UI", 11)).pack(pady=(5, 0))

feriado_frame = tk.Frame(aba_calculadora, bg="#0f172a")
feriado_inicio = tk.Entry(feriado_frame, font=("Consolas", 12), width=10, justify="center", bg="#1e293b", fg="#f8fafc")
feriado_inicio.grid(row=0, column=1, padx=5)
feriado_fim = tk.Entry(feriado_frame, font=("Consolas", 12), width=10, justify="center", bg="#1e293b", fg="#f8fafc")
feriado_fim.grid(row=0, column=3, padx=5)

adicionar_feriado_btn = tk.Button(feriado_frame, text="Adicionar Feriado", font=("Segoe UI", 10, "bold"), bg="#334155", fg="white", command=adicionar_feriado)
adicionar_feriado_btn.grid(row=0, column=4, padx=5)

# === Aba Notificações ===
aba_notificacoes = tk.Frame(abas, bg="#0f172a")
abas.add(aba_notificacoes, text="  Notificações  ", image=icon_notification, compound="left")

# === Aba Listagem ===
aba_listagem = tk.Frame(abas, bg="#0f172a")
abas.add(aba_listagem, text="  Listagem de Prazos  ", image=icon_listagem, compound="left")

# === Aba Configurações ===
aba_config = tk.Frame(abas, bg="#0f172a")
abas.add(aba_config, text="  Configurações  ", image=icon_config, compound="left")

# Rodapé
tk.Label(janela, text="Aplicação desenvolvida por: Rodrigo Junqueira de Lima Siqueira", font=("Segoe UI", 8), bg="#0f172a", fg="#64748b").pack(side="bottom", pady=(10, 5))

janela.mainloop()
