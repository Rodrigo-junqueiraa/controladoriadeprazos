"""Aba de Notificações e Histórico"""

import tkinter as tk
from tkinter import ttk, messagebox
from ...core.configuracoes import TEMA
from ..styles import criar_botao_padrao, criar_label_padrao


class AbaNotificacoes:
    """Aba de notificações do dia"""

    def __init__(self, parent, gerenciador_notif, storage):
        self.frame = tk.Frame(parent, bg=TEMA["bg_principal"])
        self.gerenciador_notif = gerenciador_notif
        self.storage = storage
        self._criar_interface()
        self._carregar_dados()

    def _criar_interface(self):
        """Cria elementos da interface"""
        # Título
        titulo = criar_label_padrao(
            self.frame,
            "Histórico de Notificações",
            font=TEMA["font_subtitulo"]
        )
        titulo.pack(pady=(20, 10))

        # Treeview
        self.colunas = ("Cliente", "Processo", "Tipo de Prazo", "Data Fatal", "Data de Registro")
        self.tree = ttk.Treeview(self.frame, columns=self.colunas, show="headings", height=20)

        for col in self.colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        self.tree.pack(fill="both", padx=10, pady=(0, 10), expand=True)

        # Botão limpar
        btn_limpar = criar_botao_padrao(
            self.frame,
            "Limpar Notificações",
            self._limpar_notificacoes,
            cor_bg=TEMA["cor_vermelho"]
        )
        btn_limpar.pack(pady=(0, 10))

    def _carregar_dados(self):
        """Carrega notificações para exibição"""
        self.tree.delete(*self.tree.get_children())
        notificados = self.gerenciador_notif.obter_notificados()

        for p in notificados:
            self.tree.insert("", "end", values=(
                p.get("cliente", ""),
                p.get("processo", ""),
                p.get("tipo_prazo", ""),
                p.get("data_para_notificar", ""),
                p.get("registro_em", "")
            ))

    def _limpar_notificacoes(self):
        """Limpa todas as notificações"""
        if messagebox.askyesno("Confirmação", "Deseja realmente limpar todas as notificações?"):
            if self.gerenciador_notif.limpar_notificados():
                self._carregar_dados()
                messagebox.showinfo("Sucesso", "Notificações removidas com sucesso.")
            else:
                messagebox.showerror("Erro", "Erro ao limpar notificações.")

    def atualizar(self):
        """Atualiza dados exibidos"""
        self._carregar_dados()

    def get_frame(self):
        """Retorna o frame da aba"""
        return self.frame
