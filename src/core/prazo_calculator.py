"""Módulo de cálculo de prazos jurídicos"""

from datetime import datetime, timedelta
import pandas as pd
from .configuracoes import PRAZOS


class CalculadoraPrazos:
    """Calcula prazos processuais de acordo com dias úteis"""

    def __init__(self):
        self.feriados = []
        self.prazos_config = PRAZOS

    def adicionar_feriados(self, data_inicio_str, data_fim_str):
        """Adiciona feriados em período (formato DD/MM)"""
        try:
            data_inicio = datetime.strptime(data_inicio_str, "%d/%m")
            data_fim = datetime.strptime(data_fim_str, "%d/%m")
            ano_atual = datetime.now().year
            data_inicio = data_inicio.replace(year=ano_atual)
            data_fim = data_fim.replace(year=ano_atual)
            
            datas = pd.date_range(data_inicio, data_fim).to_pydatetime()
            for d in datas:
                feriados_str = d.strftime("%d/%m")
                if feriados_str not in self.feriados:
                    self.feriados.append(feriados_str)
            return True
        except ValueError:
            return False

    def limpar_feriados(self):
        """Limpa todos os feriados"""
        self.feriados.clear()

    def calcular_prazo_util(self, data_publicacao_str, dias_prazo):
        """
        Calcula o termo final de um prazo em dias úteis
        
        Args:
            data_publicacao_str: Data de publicação no formato DD/MM
            dias_prazo: Número de dias úteis
            
        Returns:
            str: Data do término do prazo em formato DD/MM ou mensagem de erro
        """
        try:
            data_inicial = datetime.strptime(data_publicacao_str, "%d/%m")
            data_inicial = data_inicial.replace(year=datetime.now().year)
            
            # Gera dias úteis (seg-sex)
            df = pd.date_range(data_inicial + timedelta(days=1), periods=90, freq='B')
            
            # Filtra removendo feriados
            df_filtrado = [d for d in df if d.strftime("%d/%m") not in self.feriados]
            
            if len(df_filtrado) < dias_prazo:
                return "Prazo ultrapassa os dias úteis disponíveis"
            
            termo_final = df_filtrado[dias_prazo - 1]
            return termo_final.strftime("%d/%m")
        
        except Exception as e:
            return f"Erro: {e}"

    def calcular_com_ramo_tipo(self, data_publicacao_str, ramo, tipo_prazo):
        """Calcula prazo usando configurações de ramo e tipo"""
        if ramo not in PRAZOS:
            return "Ramo inválido"
        
        if tipo_prazo not in PRAZOS[ramo]:
            return "Tipo de prazo inválido"
        
        dias = PRAZOS[ramo][tipo_prazo]
        return self.calcular_prazo_util(data_publicacao_str, dias)

    def obter_ramos(self):
        """Retorna lista de ramos disponíveis"""
        return list(PRAZOS.keys())

    def obter_tipos_prazo(self, ramo):
        """Retorna tipos de prazo para um ramo específico"""
        return list(PRAZOS.get(ramo, {}).keys())

    def obter_dias_prazo(self, ramo, tipo):
        """Retorna número de dias de prazo"""
        return PRAZOS.get(ramo, {}).get(tipo, None)
