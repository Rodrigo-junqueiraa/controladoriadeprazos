"""Gestão de advogados memorados"""

import json
import os


class GerenciadorAdvogados:
    """Gerencia lista de advogados memorados"""

    def __init__(self, arquivo_json="advogados.json"):
        self.arquivo = arquivo_json

    def carregar(self):
        """Carrega lista de advogados"""
        if os.path.exists(self.arquivo):
            try:
                with open(self.arquivo, "r", encoding="utf-8") as f:
                    return json.load(f).get("advogados", [])
            except Exception:
                return []
        return []

    def salvar(self, lista):
        """Salva lista de advogados"""
        try:
            with open(self.arquivo, "w", encoding="utf-8") as f:
                json.dump({"advogados": lista}, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def adicionar(self, nome):
        """Adiciona novo advogado"""
        lista = self.carregar()
        if nome not in lista:
            lista.append(nome)
            return self.salvar(lista)
        return False

    def remover(self, nome):
        """Remove advogado"""
        lista = self.carregar()
        if nome in lista:
            lista.remove(nome)
            return self.salvar(lista)
        return False

    def existe(self, nome):
        """Verifica se advogado está memorado"""
        return nome in self.carregar()
