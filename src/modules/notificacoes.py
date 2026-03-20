"""Sistema de notificações de prazos"""

from datetime import datetime


class GerenciadorNotificacoes:
    """Gerencia notificações de prazos fatais"""

    def __init__(self, storage):
        """
        Args:
            storage: Instância de StorageJSON
        """
        self.storage = storage

    def verificar_prazos_hoje(self):
        """Verifica prazos com notificação para hoje"""
        hoje = datetime.now().strftime("%d/%m")
        prazos = self.storage.carregar()
        alertas = []

        for prazo in prazos:
            data_notificar = prazo.get("data_para_notificar", "").strip()
            if data_notificar == hoje:
                alertas.append(prazo)
                prazo["notificado"] = True

        if alertas:
            self.storage.salvar(prazos)

        return alertas

    def obter_formatado(self, alertas):
        """Formata alertas para exibição"""
        linhas = []
        for p in alertas:
            linha = "{} - {} - {} (FATAL)".format(
                p.get('cliente', 'N/A'),
                p.get('tipo_prazo', 'N/A'),
                p.get('processo', 'N/A')
            )
            linhas.append(linha)
        return "\n".join(linhas)

    def obter_notificados(self):
        """Retorna histórico de notificados"""
        return self.storage.filtrar_notificados()

    def obter_proximas_notificacoes(self):
        """Retorna próxima notificações (não notificadas)"""
        return self.storage.filtrar_nao_notificados()

    def limpar_notificados(self):
        """Limpa histórico de notificados"""
        return self.storage.limpar_notificados()
