"""Aba de Listagem de Prazos"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from ...core.configuracoes import TEMA
from ..styles import criar_botao_padrao, criar_label_padrao


class AbaListagem:
    """Aba para listar prazos por data"""

    def __init__(self, parent, storage):
        self.frame = tk.Frame(parent, bg=TEMA["bg_principal"])
        self.storage = storage
        self._criar_interface()

    def _criar_interface(self):
        """Cria elementos da interface"""
        # Título
        titulo = criar_label_padrao(
            self.frame,
            "Listar prazos por data",
            font=TEMA["font_subtitulo"]
        )
        titulo.pack(pady=(20, 5))

        # Frame de busca
        frame_data = tk.Frame(self.frame, bg=TEMA["bg_principal"])
        frame_data.pack(pady=5)

        label_data = criar_label_padrao(frame_data, "Data (DD/MM):", font=("Segoe UI", 10))
        label_data.pack(side="left", padx=5)

        self.entry_data_busca = tk.Entry(
            frame_data,
            font=("Segoe UI", 11),
            width=10,
            justify="center"
        )
        self.entry_data_busca.pack(side="left", padx=5)

        # Dica
        dica = criar_label_padrao(
            self.frame,
            "⚠️ Dica: Busque pela data FATAL do prazo (ex: 20/03). A busca encontra prazos registrados nesta data.",
            font=("Segoe UI", 8)
        )
        dica.config(fg="#facc15")
        dica.pack(pady=(2, 10), padx=10, anchor="w")

        # Treeview
        self.colunas = ("Cliente", "Processo", "Tipo de Prazo", "Data Fatal", "Data Notificar", "Registro em")
        self.tree = ttk.Treeview(self.frame, columns=self.colunas, show="headings", height=20)

        for col in self.colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        self.tree.pack(fill="both", padx=10, pady=(10, 5), expand=True)

        # Botão buscar
        btn_buscar = criar_botao_padrao(
            self.frame,
            "Buscar",
            self._buscar
        )
        btn_buscar.pack(pady=(0, 10))

    def _buscar(self):
        """Busca prazos pela data"""
        data = self.entry_data_busca.get().strip()

        if not data:
            messagebox.showwarning("Aviso", "Insira uma data no formato DD/MM.")
            return

        if not self._validar_data(data):
            messagebox.showerror("Erro", "Data inválida. Use o formato DD/MM.")
            return

        self.tree.delete(*self.tree.get_children())
        encontrados = self.storage.filtrar_por_data(data)

        if not encontrados:
            messagebox.showinfo("Resultado", "Nenhum prazo encontrado para essa data.")
            return

        for p in encontrados:
            self.tree.insert("", "end", values=(
                p.get("cliente", ""),
                p.get("processo", ""),
                p.get("tipo_prazo", ""),
                p.get("data_fatal", ""),
                p.get("data_para_notificar", ""),
                p.get("registro_em", "")
            ))

    def _validar_data(self, data_str):
        """Valida formato de data DD/MM"""
        try:
            datetime.strptime(data_str, "%d/%m")
            return True
        except ValueError:
            return False

    def get_frame(self):
        """Retorna o frame da aba"""
        return self.frame
