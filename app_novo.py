"""
SISTEMA DE CONTROLE DE PRAZOS JURÍDICOS
Ponto de entrada da aplicação

Desenvolvido por: Rodrigo Junqueira de Lima Siqueira
https://github.com/Rodrigo-junqueiraa
"""

import sys
import os

# Adicionar src ao caminho
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.main import main

if __name__ == "__main__":
    main()
