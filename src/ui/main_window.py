"""Janela principal da aplicação"""

import tkinter as tk
from tkinter import ttk
from ..core.configuracoes import TEMA
from ..core.utils import recurso_path
from .styles import aplicar_estilo_global
from .tabs.aba_inicio import AbaInicio
from .tabs.aba_calculadora import AbaCalculadora
from .tabs.aba_notificacoes import AbaNotificacoes
from .tabs.aba_listagem import AbaListagem
from .tabs.aba_djen import AbaSearchDJEN


class JanelaPrincipal:
    """Janela principal com système de abas"""

    def __init__(self, dependencias):
        """
        Inicializa a janela principal
        
        Args:
            dependencias: Dict com instâncias dos módulos necessários
                - calculadora: CalculadoraPrazos
                - storage: StorageJSON
                - gerenciador_notif: GerenciadorNotificações
                - gerenciador_adv: GerenciadorAdvogados
        """
        self.janela = tk.Tk()
        self.janela.title("Sistema de Controle de Prazos")
        self.janela.geometry("1300x720")
        self.janela.configure(bg=TEMA["bg_principal"])

        self.dependencias = dependencias
        self._carregar_icones()
        self._criar_interface()

    def _carregar_icones(self):
        """Carrega ícones das abas"""
        try:
            self.icon_inicio = tk.PhotoImage(file=recurso_path("icon_inicio.png")).subsample(40, 40)
            self.icon_calc = tk.PhotoImage(file=recurso_path("icon_calc.png")).subsample(40, 40)
            self.icon_notification = tk.PhotoImage(file=recurso_path("icon_notification.png")).subsample(40, 40)
            self.icon_listagem = tk.PhotoImage(file=recurso_path("icon_listagem.png")).subsample(40, 40)
            self.icon_config = tk.PhotoImage(file=recurso_path("icon_config.png")).subsample(40, 40)
        except Exception as e:
            print(f"Aviso ao carregar ícones: {e}")
            self.icon_inicio = None
            self.icon_calc = None
            self.icon_notification = None
            self.icon_listagem = None
            self.icon_config = None

    def _criar_interface(self):
        """Cria a interface com abas"""
        # Aplicar estilos globais
        aplicar_estilo_global(self.janela)

        # Notebook (abas)
        self.abas = ttk.Notebook(self.janela, style="TNotebook")
        self.abas.pack(fill="both", expand=True)

        # Aba Início
        aba_inicio = AbaInicio(self.abas)
        imagem_icon = {"image": self.icon_inicio, "compound": "left"} if self.icon_inicio else {}
        self.abas.add(aba_inicio.get_frame(), text="  Início  ", **imagem_icon)
        self.aba_inicio = aba_inicio

        # Aba Calculadora
        aba_calc = AbaCalculadora(self.abas, self.dependencias["calculadora"])
        imagem_icon = {"image": self.icon_calc, "compound": "left"} if self.icon_calc else {}
        self.abas.add(aba_calc.get_frame(), text="  Calculadora  ", **imagem_icon)
        self.aba_calculadora = aba_calc

        # Aba de Notificações
        aba_notif = AbaNotificacoes(self.abas, self.dependencias["gerenciador_notif"], self.dependencias["storage"])
        imagem_icon = {"image": self.icon_notification, "compound": "left"} if self.icon_notification else {}
        self.abas.add(aba_notif.get_frame(), text="  Notificações", **imagem_icon)
        self.aba_notificacoes = aba_notif

        # Aba de Listagem
        aba_lista = AbaListagem(self.abas, self.dependencias["storage"])
        imagem_icon = {"image": self.icon_listagem, "compound": "left"} if self.icon_listagem else {}
        self.abas.add(aba_lista.get_frame(), text="  Listagem  ", **imagem_icon)
        self.aba_listagem = aba_lista

        # Aba DJEN
        aba_djen = AbaSearchDJEN(self.abas, self.dependencias["gerenciador_adv"])
        imagem_icon = {"image": self.icon_config, "compound": "left"} if self.icon_config else {}
        self.abas.add(aba_djen.get_frame(), text="  Busca DJEN  ", **imagem_icon)
        self.aba_djen = aba_djen

    def conectar_callbacks(self, callbacks):
        """
        Conecta callbacks dos botões
        
        Args:
            callbacks: Dict com funções para cada ação
                - btn_conferir_fatais
                - btn_gerar_relatorio
                - btn_preencher
        """
        if "btn_conferir_fatais" in callbacks:
            self.aba_inicio.btn_conferir_fatais.config(command=callbacks["btn_conferir_fatais"])

        if "btn_gerar_relatorio" in callbacks:
            self.aba_inicio.btn_gerar_relatorio.config(command=callbacks["btn_gerar_relatorio"])

        if "btn_preencher" in callbacks:
            self.aba_inicio.btn_preencher.config(command=callbacks["btn_preencher"])

    def executar(self):
        """Inicia a janela principal"""
        self.janela.mainloop()

    def fechar(self):
        """Fecha a janela"""
        self.janela.quit()

    def atualizar_notificacoes(self):
        """Atualiza dados da aba de notificações"""
        if hasattr(self, 'aba_notificacoes'):
            self.aba_notificacoes.atualizar()
