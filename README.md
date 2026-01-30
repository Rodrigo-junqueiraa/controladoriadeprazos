# Controladoria de Prazos - Jurídico Digital

Aplicação com interface gráfica voltada para advogados(as) e profissionais do Direito, que automatiza o **preenchimento da data de prazo** em planilhas e realiza o cálculo do termo final com base em dias úteis, conforme a CLT e o CPC.

Versão de testes

<img width="1061" height="851" alt="image" src="https://github.com/user-attachments/assets/a5435b52-0ef0-43a7-88a6-ea6cf243e9de" />



## ✨ Funcionalidades

- 🗓 Preenchimento da data de prazo em planilha Excel
- 📅 Cálculo de prazos processuais conforme a data da publicação
- 📚 Consideração de prazos diferentes para o ramo trabalhista (CLT) e cível (CPC)
- 👨‍⚖️ Interface visual com design jurídico

## 📂 Como usar

1. Instale as dependências:

```bash
pip install pandas openpyxl pillow
```

2. Execute a aplicação:

```bash
python app.py
```

3. Para empacotar como `.exe`:

```bash
pyinstaller --noconsole --onefile --add-data "justica.png;." app.py
```

> O arquivo Excel modelo (`Planilha de prazos - atualizada.xlsx`) e a imagem `justica.png` devem estar na mesma pasta do script.

## 📌 Exemplo de prazos configurados

| Recurso                                | Ramo | Prazo (dias úteis) |
|----------------------------------------|------|---------------------|
| Embargos de declaração                 | CLT  | 5                   |
| Recurso Ordinário Trabalhista          | CLT  | 8                   |
| Recurso de Revista                     | CLT  | 8                   |
| Agravo de Petição                      | CLT  | 8                   |
| Apelação                               | CPC  | 15                  |
| Recurso Especial / Extraordinário      | CPC  | 15                  |

## 👤 Desenvolvido por

**Rodrigo Junqueira de Lima Siqueira**  
[github.com/Rodrigo-junqueiraa](https://github.com/Rodrigo-junqueiraa)
