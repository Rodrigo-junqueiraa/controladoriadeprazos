"""Geração de relatórios em PDF"""

from fpdf import FPDF
from datetime import datetime
import os
from ..core.utils import limpar_texto


class GeradorRelatorioPDF:
    """Gera relatórios em PDF de prazos fatais"""

    def __init__(self):
        self.fonte_path = "DejaVuSans.ttf"
        self.fonte_bold_path = "DejaVuSans-Bold.ttf"
        self.fonte_italic_path = "DejaVuSans-Oblique.ttf"
        self.imagem_path = "justica.png"

    def gerar_relatorio_fatais(self, prazos_filtrados, data_escolhida, caminho_saida):
        """
        Gera relatório PDF de prazos fatais
        
        Args:
            prazos_filtrados: Lista de prazos para incluir
            data_escolhida: Data no formato DD/MM
            caminho_saida: Caminho onde salvar o PDF
            
        Returns:
            bool: True se sucesso, False caso contrário
        """
        try:
            pdf = FPDF()
            pdf.add_page()

            # Configurar fontes (se existirem)
            if os.path.exists(self.fonte_path):
                pdf.add_font("DejaVu", "", self.fonte_path, uni=True)
                pdf.add_font("DejaVu", "B", self.fonte_bold_path, uni=True)
                pdf.add_font("DejaVu", "I", self.fonte_italic_path, uni=True)

            # Fundo escuro
            pdf.set_fill_color(15, 23, 42)
            pdf.rect(0, 0, 210, 297, 'F')

            # Imagem
            if os.path.exists(self.imagem_path):
                pdf.image(self.imagem_path, x=(210 - 30) / 2, y=10, w=30)

            # Título
            pdf.set_y(45)
            pdf.set_text_color(255, 255, 255)
            try:
                pdf.set_font("DejaVu", "B", 14)
            except:
                pdf.set_font("Arial", "B", 14)

            pdf.cell(0, 10, "Sistema de Controle de Prazos Jurídicos", ln=True, align="C")

            # Corpo
            pdf.set_left_margin(10)
            pdf.set_right_margin(10)
            pdf.set_auto_page_break(auto=True, margin=15)

            try:
                pdf.set_font("DejaVu", "", 12)
            except:
                pdf.set_font("Arial", "", 12)

            pdf.ln(10)
            pdf.cell(0, 10, f"Data do Relatório: {datetime.now().strftime('%d/%m/%Y')}", ln=True)
            pdf.cell(0, 10, f"Prazos Fatais - {data_escolhida}", ln=True)
            pdf.ln(5)

            # Prazos
            for p in prazos_filtrados:
                try:
                    cliente = p.get("cliente", "N/A")
                    processo = p.get("processo", "N/A")
                    tipo = p.get("tipo_prazo", "N/A")

                    linha = limpar_texto(f"{cliente} - {processo} - {tipo}")

                    if not linha:
                        linha = "Dados indisponíveis"

                    try:
                        pdf.set_font("DejaVu", "", 12)
                    except:
                        pdf.set_font("Arial", "", 12)

                    pdf.multi_cell(180, 10, linha)
                    pdf.ln(2)

                except Exception:
                    try:
                        pdf.set_font("DejaVu", "", 12)
                    except:
                        pdf.set_font("Arial", "", 12)
                    pdf.cell(0, 8, "[Erro ao exibir linha]", ln=True)
                    pdf.ln(3)

            # Rodapé
            pdf.ln(10)
            try:
                pdf.set_font("DejaVu", "I", 10)
            except:
                pdf.set_font("Arial", "I", 10)

            pdf.set_text_color(200, 200, 200)
            pdf.cell(0, 10, "Relatório gerado automaticamente pelo sistema.", ln=True, align="C")
            pdf.cell(0, 10, "Desenvolvido por Rodrigo Junqueira de Lima Siqueira", ln=True, align="C")

            # Salvar
            pdf.output(caminho_saida)
            return True

        except Exception as e:
            print(f"Erro ao gerar PDF: {e}")
            return False
