# Estrutura Nova - Sistema de Controle de Prazos

## 📁 Organização de Pastas

```
src/
├── core/                       # Lógica central e cálculos
│   ├── configuracoes.py       # Constantes (PRAZOS, TEMA, ARQUIVOS)
│   ├── prazo_calculator.py    # CalculadoraPrazos (cálculos de prazos)
│   └── utils.py               # Funções utilitárias (limpar_texto, validações)
│
├── data/                       # Persistência (JSON, Excel)
│   ├── storage.py             # StorageJSON (gerencia dados em JSON)
│   └── excel_handler.py       # ExcelHandler (lê/escreve Excel)
│
├── modules/                    # Funcionalidades especializadas
│   ├── advogados.py           # GerenciadorAdvogados (CRUD advogados)
│   ├── notificacoes.py        # GerenciadorNotificacoes (alertas de prazos)
│   └── relatorio_pdf.py       # GeradorRelatorioPDF (gera PDFs)
│
├── ui/                         # Interface gráfica (Tkinter)
│   ├── styles.py              # Estilos globais e temas
│   ├── main_window.py         # JanelaPrincipal (gerencia abas)
│   └── tabs/                  # Cada aba em arquivo separado
│       ├── aba_inicio.py      # Aba inicial
│       ├── aba_calculadora.py # Calculadora de prazos
│       ├── aba_notificacoes.py# Histórico de notificações
│       ├── aba_listagem.py    # Busca de prazos por data
│       └── aba_djen.py        # Gestão de advogados e busca DJEN
│
└── main.py                     # AplicacaoPrincipal (integra tudo)
```

## 🔑 Classes Principais

| Classe                    | Arquivo                    | Função                       |
| ------------------------- | -------------------------- | ---------------------------- |
| `CalculadoraPrazos`       | `core/prazo_calculator.py` | Calcula prazos em dias úteis |
| `StorageJSON`             | `data/storage.py`          | Persiste dados em JSON       |
| `ExcelHandler`            | `data/excel_handler.py`    | Preenche planilhas Excel     |
| `GerenciadorAdvogados`    | `modules/advogados.py`     | Memoriza advogados           |
| `GerenciadorNotificacoes` | `modules/notificacoes.py`  | Gerencia alertas             |
| `GeradorRelatorioPDF`     | `modules/relatorio_pdf.py` | Gera PDFs                    |
| `JanelaPrincipal`         | `ui/main_window.py`        | Gerencia interface           |

## 🔗 Fluxo de Integração

```
app_novo.py (entrada)
    ↓
src/main.py (AplicacaoPrincipal)
    ↓
    ├→ Módulos Core
    │  ├→ CalculadoraPrazos
    │  ├→ StorageJSON
    │  └─→ ExcelHandler
    │
    ├→ Módulos Especializados
    │  ├→ GerenciadorAdvogados
    │  ├→ GerenciadorNotificacoes
    │  └─→ GeradorRelatorioPDF
    │
    └→ UI (JanelaPrincipal)
       ├→ AbaInicio
       ├→ AbaCalculadora
       ├→ AbaNotificacoes
       ├→ AbaListagem
       └─→ AbaSearchDJEN
```

## 💡 Exemplos de Uso

### Calcular Prazo

```python
from src.core.prazo_calculator import CalculadoraPrazos

calc = CalculadoraPrazos()
resultado = calc.calcular_prazo_util("29/04", 8)
print(resultado)  # Retorna: "13/05" (por exemplo)
```

### Salvar Prazo em JSON

```python
from src.data.storage import StorageJSON

storage = StorageJSON("CAMINHO_JSON")
prazo = {
    "cliente": "João Silva",
    "processo": "0001234-56.2024",
    "data_fatal": "13/05",
    "notificado": False
}
storage.adicionar_prazo(prazo)
```

### Gerar PDF

```python
from src.modules.relatorio_pdf import GeradorRelatorioPDF

gerador = GeradorRelatorioPDF()
prazos = [{"cliente": "João", "processo": "001", "tipo_prazo": "RR"}]
gerador.gerar_relatorio_fatais(prazos, "13/05", "relatorio.pdf")
```

## ✨ Vantagens

- **Modular**: Cada função em seu próprio lugar
- **Testável**: Componentes independentes
- **Escalável**: Fácil adicionar novas funcionalidades
- **Mantível**: Código organizado e comentado
- **Reutilizável**: Módulos podem funcionar isoladamente

## 🚀 Próximas Etapas

1. ✅ Executar teste básico: `python app_novo.py`
2. ⬜ Implementar testes unitários
3. ⬜ Empacotar como `.exe` com PyInstaller
4. ⬜ Adicionar mais módulos conforme necessário
