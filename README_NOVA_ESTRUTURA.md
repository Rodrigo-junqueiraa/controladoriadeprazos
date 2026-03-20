```
ESTRUTURA REORGANIZADA - Sistema de Controle de Prazos Jurídicos
________________________________________________________________

Nova Arquitetura
================

Antes:  Tudo em um único arquivo (app.py)
Depois: Código organizado em módulos funcionais independentes

Estrutura de Pastas:
====================

preencher_pz/
├── src/
│   ├── __init__.py
│   ├── main.py                    # Integração principal
│   │
│   ├── core/                      # Lógica central
│   │   ├── __init__.py
│   │   ├── configuracoes.py       # Constantes e temas
│   │   ├── utils.py               # Funções utilitárias
│   │   └── prazo_calculator.py    # Classe CalculadoraPrazos
│   │
│   ├── data/                      # Persistência de dados
│   │   ├── __init__.py
│   │   ├── storage.py             # Classe StorageJSON
│   │   └── excel_handler.py       # Classe ExcelHandler
│   │
│   ├── modules/                   # Funcionalidades especializadas
│   │   ├── __init__.py
│   │   ├── advogados.py           # Classe GerenciadorAdvogados
│   │   ├── notificacoes.py        # Classe GerenciadorNotificacoes
│   │   └── relatorio_pdf.py       # Classe GeradorRelatorioPDF
│   │
│   └── ui/                        # Interface gráfica
│       ├── __init__.py
│       ├── styles.py              # Estilos e temas
│       ├── main_window.py         # Classe JanelaPrincipal
│       └── tabs/                  # Abas da interface
│           ├── __init__.py
│           ├── aba_inicio.py      # Aba de início
│           ├── aba_calculadora.py # Aba calculadora
│           ├── aba_notificacoes.py# Aba notificações
│           ├── aba_listagem.py    # Aba listagem
│           └── aba_djen.py        # Aba busca DJEN
│
├── app_novo.py                    # Ponto de entrada (substitui app.py)
├── requirements.txt               # Dependências Python
└── README_NOVA_ESTRUTURA.md      # Este arquivo


Benefícios da Nova Estrutura
=============================

✅ Separação de Responsabilidades
   - Cada módulo tem uma função específica
   - Fácil manutenção e testes

✅ Modularidade
   - Componentes independentes e reutilizáveis
   - Fácil adicionar novas funcionalidades

✅ Escalabilidade
   - Código organizado para crescimento futuro
   - Menos conflitos entre funcionalidades

✅ Reusabilidade
   - Módulos podem ser usados em outros projetos
   - Testes unitários simplificados

✅ Manutenibilidade
   - Código mais legível e compreensível
   - Menos duplicação


Mapeamento de Funcionalidades Original → Nova Estrutura
========================================================

CALCULADORA DE PRAZOS
  Original: app.py (linhas ~700-800)
  Novo:     src/core/prazo_calculator.py (Classe CalculadoraPrazos)

GESTÃO DE JSON
  Original: app.py (funções carregar_prazos, salvar_prazos, etc)
  Novo:     src/data/storage.py (Classe StorageJSON)

GESTÃO DE EXCEL
  Original: app.py (função gerar_excel)
  Novo:     src/data/excel_handler.py (Classe ExcelHandler)

GESTÃO DE ADVOGADOS
  Original: app.py (funções carregar_advogados, salvar_advogados)
  Novo:     src/modules/advogados.py (Classe GerenciadorAdvogados)

NOTIFICAÇÕES
  Original: app.py (verificar_prazos_hoje, carregar_notificacoes)
  Novo:     src/modules/notificacoes.py (Classe GerenciadorNotificacoes)

RELATÓRIOS PDF
  Original: app.py (função gerar_relatorio_fatais_pdf)
  Novo:     src/modules/relatorio_pdf.py (Classe GeradorRelatorioPDF)

INTERFACE GRÁFICA
  Original: app.py (interface Tkinter toda integrada)
  Novo:     src/ui/
            - main_window.py (Janela principal)
            - tabs/aba_*.py (Cada aba em arquivo separado)
            - styles.py (Estilos e temas)


Como Usar o Novo Sistema
========================

1. Executar a aplicação:
   python app_novo.py

2. Estrutura de imports dentro do projeto:
   # No módulo UI
   from ..core.prazo_calculator import CalculadoraPrazos
   from ..data.storage import StorageJSON
   
   # No módulo principal
   from .core.configuracoes import PRAZOS
   from .modules.notificacoes import GerenciadorNotificacoes

3. Adicionar nova funcionalidade:
   - Criar novo arquivo em src/modules/
   - Importar em src/main.py
   - Conectar à UI conforme necessário


Próximas Melhorias
=================

□ Implementar aba de Notificações Futuras completa
□ Interface de preenchimento de múltiplos prazos
□ Testes unitários para cada módulo
□ Arquivo requirements.txt atualizado
□ Logging centralizado
□ Configurações em arquivo externo
□ Banco de dados (SQLite) opcional


Notas Técnicas
==============

- Todas as classes seguem padrão OOP
- Dependency injection usada em JanelaPrincipal
- Factory functions para elementos UI repetitivosRRRR
- Tratamento de exceções em pontos críticos
- Imports organizados por funcionalidade


Migrando do Código Antigo
=========================

Se você tem customizações no app.py original:

1. Identifique em qual submódulo a funcionalidade pertence
2. Encontre a classe correspondente
3. Adicione método ou propriedade conforme necessário
4. Importe e use em src/main.py
5. Conecte callbacks na UI conforme necessário

Exemplo:
  app.py: def minha_funcao() { ... }
  Novo:   src/modules/custom.py: class MinhaClasse { def minha_funcao(self) {...} }
  main.py: self.minha_instancia = MinhaClasse()
  UI:      callbacks["meu_botao"] = self.minha_instancia.minha_funcao


Problemas Comuns & Soluções
===========================

Erro: "ModuleNotFoundError: No module named 'src'"
Solução: Executar sempre de dentro da pasta raiz (preencher_pz)

Erro: "Arquivo não encontrado justica.png"
Solução: Verificar se arquivo recursos está na raiz do projeto

Erro: "JSON não encontrado"
Solução: StorageJSON cria arquivo automaticamente na primeira execução


Dúvidas?
========

Consulte a documentação dos módulos:
- Cada arquivo .py tem docstrings detalhadas
- Cada classe tem comentários explicativos
- Use help(NomeClasse) no Python para mais info
```
