"""Estilos e temas globais da interface"""

import tkinter as tk
from tkinter import ttk
from ..core.configuracoes import TEMA


def aplicar_estilo_global(janela):
    """Aplica estilos globais à aplicação"""
    
    # Estilo do notebook
    style = ttk.Style()
    style.theme_use('default')
    style.configure("TNotebook", 
                    background=TEMA["bg_principal"], 
                    borderwidth=0)
    style.configure("TNotebook.Tab",
                    background=TEMA["bg_secundario"],
                    foreground=TEMA["fg_principal"],
                    font=("Segoe UI", 10, "bold"),
                    padding=(20, 10))
    style.map("TNotebook.Tab",
              background=[("selected", TEMA["cor_azul"])],
              foreground=[("selected", TEMA["fg_principal"])])
    
    # Estilo do Treeview
    style.configure("Treeview",
                    background=TEMA["bg_secundario"],
                    foreground=TEMA["fg_principal"],
                    fieldbackground=TEMA["bg_secundario"],
                    font=("Segoe UI", 9))
    style.configure("Treeview.Heading",
                    background=TEMA["bg_principal"],
                    foreground=TEMA["fg_principal"],
                    font=("Segoe UI", 9, "bold"))
    style.map("Treeview",
              background=[("selected", TEMA["cor_azul"])],
              foreground=[("selected", TEMA["fg_principal"])])


def configurar_estilo_botoes(widget):
    """Aplica estilos padrão a botões"""
    widget.configure(
        font=("Segoe UI", 9, "bold"),
        padx=4,
        pady=2,
        relief="raised",
        bd=2
    )


def efeito_hover_botao(event):
    """Efeito hover ao passar mouse sobre botão"""
    event.widget.config(bg=TEMA["cor_hover"])


def efeito_sair_hover_botao(event):
    """Remove efeito hover ao sair do botão"""
    event.widget.config(bg=event.widget.original_bg)


def criar_botao_padrao(parent, texto, comando, cor_bg=None, **kwargs):
    """Factory function para criar botões padrão"""
    if cor_bg is None:
        cor_bg = TEMA["cor_azul"]
    
    # Remove font de kwargs se foi passado, para evitar conflito
    fonte = kwargs.pop('font', ("Segoe UI", 10, "bold"))
    
    btn = tk.Button(
        parent,
        text=texto,
        command=comando,
        bg=cor_bg,
        fg=TEMA["fg_principal"],
        font=fonte,
        **kwargs
    )
    
    btn.original_bg = cor_bg
    configurar_estilo_botoes(btn)
    btn.bind("<Enter>", efeito_hover_botao)
    btn.bind("<Leave>", efeito_sair_hover_botao)
    
    return btn


def criar_label_padrao(parent, texto, **kwargs):
    """Factory function para criar labels padrão"""
    return tk.Label(
        parent,
        text=texto,
        bg=TEMA["bg_principal"],
        fg=TEMA["fg_principal"],
        **kwargs
    )
