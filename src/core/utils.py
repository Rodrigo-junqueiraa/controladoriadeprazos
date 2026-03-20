"""Utilitários gerais do sistema"""

import os
import sys
import unicodedata


def recurso_path(rel_path):
    """Retorna caminho correto de recursos (dev/exe)"""
    try:
        base_path = sys._MEIPASS  # PyInstaller
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)


def caminho_json_prazos():
    """Retorna caminho do JSON de prazos"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        # Caminho da raiz do projeto (2 níveis acima de src/core/)
        base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, "CAMINHO_JSON")


def limpar_texto(texto):
    """Remove caracteres especiais e inválidos do texto"""
    texto = texto.replace("–", "-")
    texto = ''.join(c for c in texto if unicodedata.category(c)[0] != "C")
    texto = ''.join(c if c.isprintable() else "?" for c in texto)
    return texto.strip()


def validar_formato_data(data_str):
    """Valida formato DD/MM"""
    from datetime import datetime
    try:
        datetime.strptime(data_str, "%d/%m")
        return True
    except ValueError:
        return False
