"""
Arquivo principal de integração
Inicializa todos os módulos e executa a aplicação
"""

import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from datetime import datetime
import os

# Importações de módulos
from .core.prazo_calculator import CalculadoraPrazos
from .core.utils import validar_formato_data, caminho_json_prazos
from .core.configuracoes import ARQUIVOS
from .data.storage import StorageJSON
from .data.excel_handler import ExcelHandler
from .modules.advogados import GerenciadorAdvogados
from .modules.notificacoes import GerenciadorNotificacoes
from .modules.relatorio_pdf import GeradorRelatorioPDF
from .ui.main_window import JanelaPrincipal
from .ui.janela_preenchimento import JanelaPreenchimento


class AplicacaoPrincipal:
    """Classe principal que integra toda a aplicação"""

    def __init__(self):
        """Inicializa a aplicação"""
        self._inicializar_modulos()
        self._criar_janela()
        self._conectar_callbacks()

    def _inicializar_modulos(self):
        """Inicializa todos os módulos de negócio"""
        print("Inicializando módulos...")

        # Calculadora de prazos
        self.calculadora = CalculadoraPrazos()

        # Storage JSON
        self.storage = StorageJSON(caminho_json_prazos())

        # Gerenciadores
        self.gerenciador_notif = GerenciadorNotificacoes(self.storage)
        self.gerenciador_adv = GerenciadorAdvogados()

        # Excel handler
        self.excel_handler = ExcelHandler(ARQUIVOS["excel_modelo"])

        # PDF
        self.gerador_pdf = GeradorRelatorioPDF()

        print("Módulos inicializados com sucesso!")

    def _criar_janela(self):
        """Cria a interface gráfica"""
        print("Criando interface gráfica...")

        dependencias = {
            "calculadora": self.calculadora,
            "storage": self.storage,
            "gerenciador_notif": self.gerenciador_notif,
            "gerenciador_adv": self.gerenciador_adv,
        }

        self.janela_principal = JanelaPrincipal(dependencias)
        print("Interface gráfica criada!")

    def _conectar_callbacks(self):
        """Conecta botões a suas funcionalidades"""
        callbacks = {
            "btn_conferir_fatais": self._conferir_prazos_fatais,
            "btn_gerar_relatorio": self._gerar_relatorio_fatais,
            "btn_preencher": self._abrir_preenchimento,
        }

        self.janela_principal.conectar_callbacks(callbacks)

    def _conferir_prazos_fatais(self):
        """Verifica prazos fatais de hoje"""
        alertas = self.gerenciador_notif.verificar_prazos_hoje()

        if alertas:
            texto = self.gerenciador_notif.obter_formatado(alertas)
            messagebox.showinfo("🔔 Notificação", f"Prazos configurados para hoje:\n\n{texto}")
        else:
            messagebox.showinfo("Notificação", "Nenhum prazo fatal para hoje.")

        self.janela_principal.atualizar_notificacoes()

    def _gerar_relatorio_fatais(self):
        """Gera relatório PDF de prazos fatais"""
        data_escolhida = simpledialog.askstring("Data", "Informe a data (DD/MM):")

        if not data_escolhida:
            return

        if not validar_formato_data(data_escolhida):
            messagebox.showerror("Erro", "Data inválida. Use o formato DD/MM.")
            return

        prazos_filtrados = self.storage.filtrar_por_data(data_escolhida)
        prazos_filtrados = [p for p in prazos_filtrados if p.get("lembrete", "").upper() == "FATAL"]

        if not prazos_filtrados:
            messagebox.showinfo("Sem prazos", "Nenhum prazo fatal encontrado para essa data.")
            return

        # Selecionar local de salvamento
        nome_arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"Relatorio_Fatais_{data_escolhida.replace('/', '-')}.pdf",
            title="Salvar Relatório PDF"
        )

        if not nome_arquivo:
            messagebox.showinfo("Cancelado", "Salvamento foi cancelado.")
            return

        # Gerar relatório
        sucesso = self.gerador_pdf.gerar_relatorio_fatais(
            prazos_filtrados,
            data_escolhida,
            nome_arquivo
        )

        if sucesso:
            messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{nome_arquivo}")
        else:
            messagebox.showerror("Erro", "Erro ao gerar PDF.")

    def _abrir_preenchimento(self):
        """Abre janela para preenchimento de múltiplos prazos"""
        JanelaPreenchimento(
            self.janela_principal.janela,
            self.excel_handler,
            self.storage
        )

    def executar(self):
        """Inicia a aplicação"""
        print("Iniciando aplicação...")
        self.janela_principal.executar()


def main():
    """Função de entrada"""
    try:
        app = AplicacaoPrincipal()
        app.executar()
    except Exception as e:
        print(f"Erro na aplicação: {e}")
        import traceback
        traceback.print_exc()
        messagebox.showerror("Erro Fatal", f"Erro na aplicação:\n{e}")


if __name__ == "__main__":
    main()
