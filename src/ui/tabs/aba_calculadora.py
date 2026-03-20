"""Aba Calculadora de Prazos"""

import tkinter as tk
from tkinter import ttk, messagebox
from ...core.utils import recurso_path
from ...core.configuracoes import TEMA
from ..styles import criar_botao_padrao, criar_label_padrao


class AbaCalculadora:
    """Aba para calcular prazos"""

    def __init__(self, parent, calculadora):
        self.frame = tk.Frame(parent, bg=TEMA["bg_principal"])
        self.calculadora = calculadora
        self._criar_interface()

    def _criar_interface(self):
        """Cria elementos da interface"""
        # Título
        titulo = criar_label_padrao(
            self.frame,
            "Digite aqui a data de publicação",
            font=TEMA["font_titulo"]
        )
        titulo.pack(pady=(100, 5), anchor="center")

        # Campo de data
        self.publicacao_entry = tk.Entry(
            self.frame,
            font=("Consolas", 14),
            width=20,
            justify="center",
            bd=2,
            relief="flat",
            bg=TEMA["bg_secundario"],
            fg="#f8fafc",
            insertbackground="#f8fafc"
        )
        self.publicacao_entry.pack(pady=5, anchor="center")

        # Seleção de ramo
        self.ramo_var = tk.StringVar()
        ramo_menu = ttk.Combobox(
            self.frame,
            textvariable=self.ramo_var,
            values=self.calculadora.obter_ramos(),
            state="readonly",
            width=40
        )
        ramo_menu.pack(pady=2, anchor="center")
        ramo_menu.bind("<<ComboboxSelected>>", self._atualizar_tipos)

        # Seleção de tipo
        self.tipo_var = tk.StringVar()
        self.tipo_menu = ttk.Combobox(
            self.frame,
            textvariable=self.tipo_var,
            state="readonly",
            width=40
        )
        self.tipo_menu.pack(pady=2, anchor="center")

        # Botão calcular
        btn_calcular = criar_botao_padrao(
            self.frame,
            "Calcular Prazo Jurídico",
            self._calcular
        )
        btn_calcular.pack(pady=10, anchor="center")

        # Checkbox feriados
        self.checkbox_var = tk.BooleanVar()
        checkbox = tk.Checkbutton(
            self.frame,
            text="Ao longo do prazo existem feriados?",
            variable=self.checkbox_var,
            command=self._alternar_feriado,
            bg=TEMA["bg_principal"],
            fg=TEMA["fg_principal"],
            selectcolor=TEMA["bg_principal"],
            font=("Segoe UI", 11)
        )
        checkbox.pack(pady=(5, 0), anchor="center")

        # Frame de feriados
        self.feriado_frame = tk.Frame(self.frame, bg=TEMA["bg_principal"])

        tk.Label(self.feriado_frame, text="de", bg=TEMA["bg_principal"], fg=TEMA["fg_principal"]).grid(row=0, column=0, padx=5)

        self.feriado_inicio = tk.Entry(
            self.feriado_frame,
            font=("Consolas", 12),
            width=10,
            justify="center",
            bg=TEMA["bg_secundario"],
            fg="#f8fafc"
        )
        self.feriado_inicio.grid(row=0, column=1, padx=5)

        tk.Label(self.feriado_frame, text="até", bg=TEMA["bg_principal"], fg=TEMA["fg_principal"]).grid(row=0, column=2, padx=5)

        self.feriado_fim = tk.Entry(
            self.feriado_frame,
            font=("Consolas", 12),
            width=10,
            justify="center",
            bg=TEMA["bg_secundario"],
            fg="#f8fafc"
        )
        self.feriado_fim.grid(row=0, column=3, padx=5)

        btn_adicionar_feriado = criar_botao_padrao(
            self.feriado_frame,
            "Adicionar Feriado",
            self._adicionar_feriado
        )
        btn_adicionar_feriado.grid(row=0, column=4, padx=5)

    def _atualizar_tipos(self, event):
        """Atualiza tipos de prazo baseado no ramo"""
        ramo = self.ramo_var.get()
        tipos = self.calculadora.obter_tipos_prazo(ramo)
        self.tipo_menu['values'] = sorted(tipos)
        self.tipo_var.set("")

    def _calcular(self):
        """Calcula o prazo"""
        data = self.publicacao_entry.get()
        ramo = self.ramo_var.get()
        tipo = self.tipo_var.get()

        if not data or not ramo or not tipo:
            messagebox.showwarning("Aviso", "Preencha todos os campos.")
            return

        resultado = self.calculadora.calcular_com_ramo_tipo(data, ramo, tipo)
        messagebox.showinfo("Resultado do Cálculo", f"{tipo}\nPrazo final: {resultado}")

    def _alternar_feriado(self):
        """Mostra/esconde frame de feriados"""
        if self.checkbox_var.get():
            self.feriado_frame.pack(pady=5)
        else:
            self.feriado_frame.pack_forget()
            self.calculadora.limpar_feriados()

    def _adicionar_feriado(self):
        """Adiciona feriado ao cálculo"""
        inicio = self.feriado_inicio.get()
        fim = self.feriado_fim.get()

        if not self.calculadora.adicionar_feriados(inicio, fim):
            messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM.")
            return

        self.feriado_inicio.delete(0, tk.END)
        self.feriado_fim.delete(0, tk.END)
        messagebox.showinfo("Feriado adicionado", f"Feriado(s) registrado(s): {inicio} até {fim}")

    def get_frame(self):
        """Retorna o frame da aba"""
        return self.frame
