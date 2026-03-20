"""Gestão de armazenamento em JSON"""

import json
import os
from datetime import datetime


class StorageJSON:
    """Gerencia persistência de dados em JSON"""

    def __init__(self, caminho_arquivo):
        self.caminho = caminho_arquivo

    def carregar(self):
        """Carrega dados do JSON"""
        if not os.path.exists(self.caminho):
            return []
        try:
            with open(self.caminho, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Erro ao carregar JSON: {e}")
            return []

    def salvar(self, dados):
        """Salva dados no JSON"""
        try:
            with open(self.caminho, "w", encoding="utf-8") as f:
                json.dump(dados, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Erro ao salvar JSON: {e}")
            return False

    def adicionar_prazo(self, prazo_dict):
        """Adiciona um novo prazo ao JSON"""
        dados = self.carregar()
        dados.append(prazo_dict)
        return self.salvar(dados)

    def atualizar_prazo(self, index, prazo_dict):
        """Atualiza um prazo existente"""
        dados = self.carregar()
        if 0 <= index < len(dados):
            dados[index] = prazo_dict
            return self.salvar(dados)
        return False

    def remover_prazo(self, index):
        """Remove um prazo pelo índice"""
        dados = self.carregar()
        if 0 <= index < len(dados):
            dados.pop(index)
            return self.salvar(dados)
        return False

    def filtrar_por_data(self, data_str):
        """Filtra prazos por data (DD/MM)"""
        dados = self.carregar()
        return [p for p in dados if p.get("data_fatal") == data_str]

    def filtrar_notificados(self):
        """Retorna apenas prazos notificados"""
        dados = self.carregar()
        return [p for p in dados if p.get("notificado")]

    def filtrar_nao_notificados(self):
        """Retorna apenas prazos não notificados"""
        dados = self.carregar()
        return [p for p in dados if not p.get("notificado")]

    def marcar_como_notificado(self, indice):
        """Marca um prazo como notificado"""
        dados = self.carregar()
        if 0 <= indice < len(dados):
            dados[indice]["notificado"] = True
            return self.salvar(dados)
        return False

    def limpar_notificados(self):
        """Remove todos os prazos notificados"""
        dados = self.carregar()
        dados_filtrados = [p for p in dados if not p.get("notificado")]
        return self.salvar(dados_filtrados)
