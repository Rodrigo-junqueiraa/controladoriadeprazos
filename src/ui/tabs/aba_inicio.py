"""Aba Principal - Início da aplicação"""

import tkinter as tk
import os
from PIL import Image, ImageTk
from ...core.utils import recurso_path
from ...core.configuracoes import TEMA
from ..styles import criar_botao_padrao, criar_label_padrao


class AbaInicio:
    """Aba principal de início"""

    def __init__(self, parent):
        self.frame = tk.Frame(parent, bg=TEMA["bg_principal"])
        self._criar_interface()

    def _criar_interface(self):
        """Cria elementos da interface"""
        # Imagem
        self._adicionar_imagem()

        # Título
        titulo = criar_label_padrao(
            self.frame,
            "Sistema de Controle de Prazos",
            font=TEMA["font_titulo"]
        )
        titulo.pack(pady=(10, 5), anchor="center")

        # Botões de ação
        self.btn_conferir_fatais = criar_botao_padrao(
            self.frame,
            "Conferir prazos fatais do dia",
            None,
            cor_bg=TEMA["cor_vermelho"]
        )
        self.btn_conferir_fatais.pack(pady=(10, 5), anchor="center")

        self.btn_gerar_relatorio = criar_botao_padrao(
            self.frame,
            "Gerar Relatório de Prazos Fatais (PDF)",
            None,
            cor_bg=TEMA["cor_vermelho"]
        )
        self.btn_gerar_relatorio.pack(pady=(5, 30), anchor="center")

        self.btn_preencher = criar_botao_padrao(
            self.frame,
            "Preencher Cliente / Processo",
            None,
            cor_bg=TEMA["cor_verde"]
        )
        self.btn_preencher.pack(pady=(10, 30), anchor="center")

    def _adicionar_imagem(self):
        """Adiciona imagem do sistema"""
        imagem_path = recurso_path("justica.png")
        if os.path.exists(imagem_path):
            try:
                imagem = Image.open(imagem_path)
                imagem = imagem.resize((140, 140), Image.LANCZOS)
                imagem_tk = ImageTk.PhotoImage(imagem)

                img_label = tk.Label(self.frame, image=imagem_tk, bg=TEMA["bg_principal"])
                img_label.image = imagem_tk
                img_label.pack(pady=(80, 5), anchor="center")
            except Exception as e:
                print(f"Erro ao carregar imagem: {e}")

    def get_frame(self):
        """Retorna o frame da aba"""
        return self.frame
