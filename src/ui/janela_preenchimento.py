"""Janela de preenchimento de múltiplos prazos"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from ..core.configuracoes import TEMA
from ..ui.styles import criar_botao_padrao, criar_label_padrao


class JanelaPreenchimento:
    """Janela para preenchimento de múltiplos prazos e geração de Excel"""

    def __init__(self, parent, excel_handler, storage):
        """
        Args:
            parent: Janela pai (root)
            excel_handler: Instância de ExcelHandler
            storage: Instância de StorageJSON
        """
        self.excel_handler = excel_handler
        self.storage = storage
        self.registros = []
        self.dado_editando = None

        self.janela = tk.Toplevel(parent)
        self.janela.title("Múltiplos Preenchimentos")
        self.janela.geometry("520x720")
        self.janela.configure(bg=TEMA["bg_principal"])

        self._criar_interface()

    def _criar_interface(self):
        """Cria elementos da interface"""
        # Frame de entrada de dados
        frame_entrada = tk.Frame(self.janela, bg=TEMA["bg_principal"])
        frame_entrada.pack(pady=5, padx=10, fill="x")

        # Campos de entrada
        campos = [
            ("Data do Prazo (DD/MM):", "entry_data_prazo"),
            ("Nome:", "entry_nome"),
            ("Processo:", "entry_processo"),
            ("Tipo de Prazo:", "entry_tipo"),
            ("Responsável:", "entry_resp"),
            ("Data para notificar (DD/MM):", "entry_data_notificar"),
            ("Publicação:", "entry_pub"),
        ]

        for label_text, attr_name in campos:
            criar_label_padrao(frame_entrada, label_text).pack()
            entry = tk.Entry(frame_entrada)
            entry.pack(fill="x", pady=2)
            setattr(self, attr_name, entry)

        # Campo de observações (Text)
        criar_label_padrao(frame_entrada, "Observações:").pack()
        self.entry_obs = tk.Text(frame_entrada, height=3, wrap="word", bg=TEMA["bg_secundario"], fg="white")
        self.entry_obs.pack(fill="x", padx=0, pady=5)

        # Treeview com registros
        colunas = ("Nome", "Processo", "Tipo", "Data", "Resp", "Pub", "Obs")
        self.tree = ttk.Treeview(self.janela, columns=colunas, show="headings", height=12)

        for col in colunas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80, anchor="center")

        self.tree.pack(pady=5, fill="both", expand=True, padx=10)

        # Frame de botões de ação
        frame_botoes = tk.Frame(self.janela, bg=TEMA["bg_principal"])
        frame_botoes.pack(pady=5)

        btn_add = criar_botao_padrao(frame_botoes, "Adicionar", self._adicionar)
        btn_add.pack(side="left", padx=3)

        btn_editar = criar_botao_padrao(frame_botoes, "Editar", self._editar)
        btn_editar.pack(side="left", padx=3)

        btn_salvar = criar_botao_padrao(frame_botoes, "Salvar Edição", self._salvar_alteracoes)
        btn_salvar.pack(side="left", padx=3)

        btn_excluir = criar_botao_padrao(frame_botoes, "Excluir", self._excluir, cor_bg=TEMA["cor_vermelho"])
        btn_excluir.pack(side="left", padx=3)

        # Botão gerar Excel
        btn_gerar = criar_botao_padrao(
            self.janela,
            "Preencher Dados Cliente/Processo (Excel)",
            self._gerar_excel,
            cor_bg=TEMA["cor_verde"]
        )
        btn_gerar.pack(pady=10)

    def _adicionar(self):
        """Adiciona novo registro"""
        nome = self.entry_nome.get().strip()
        processo = self.entry_processo.get().strip()
        tipo = self.entry_tipo.get().strip()
        data = self.entry_data_prazo.get().strip()
        resp = self.entry_resp.get().strip()
        pub = self.entry_pub.get().strip()
        obs = self.entry_obs.get("1.0", tk.END).strip()
        data_notificar = self.entry_data_notificar.get().strip()

        if not all([nome, processo, tipo, data, data_notificar, resp, pub]):
            messagebox.showwarning("Aviso", "Preencha todos os campos obrigatórios!")
            return

        # Verificar duplicidade
        for r in self.registros:
            if r[0] == nome and r[1] == processo and r[2] == tipo:
                messagebox.showwarning("Duplicidade", "Este registro já existe!")
                return

        # Adicionar registro trama para a grid
        prazo = {
            "cliente": nome,
            "processo": processo,
            "tipo_prazo": tipo,
            "data_fatal": data,
            "data_para_notificar": data_notificar,
            "lembrete": "FATAL",
            "notificado": False,
            "registro_em": datetime.now().strftime("%Y-%m-%d %H:%M")
        }

        # Salvar imediatamente no JSON
        self.storage.adicionar_prazo(prazo)

        self.registros.append((nome, processo, tipo, data, data_notificar, resp, pub, obs))
        obs_display = obs[:30] + ("..." if len(obs) > 30 else "")
        self.tree.insert("", "end", values=(nome, processo, tipo, data, resp, pub, obs_display))

        # Limpar campos
        self._limpar_campos()
        messagebox.showinfo("Sucesso", "Registro adicionado e salvo no JSON!")

    def _editar(self):
        """Carrega registro para edição"""
        item = self.tree.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um registro para editar.")
            return

        self.dado_editando = item[0]
        valores = self.tree.item(self.dado_editando, "values")

        self.entry_nome.delete(0, tk.END)
        self.entry_nome.insert(0, valores[0])
        self.entry_processo.delete(0, tk.END)
        self.entry_processo.insert(0, valores[1])
        self.entry_tipo.delete(0, tk.END)
        self.entry_tipo.insert(0, valores[2])
        self.entry_data_prazo.delete(0, tk.END)
        self.entry_data_prazo.insert(0, valores[3])
        self.entry_resp.delete(0, tk.END)
        self.entry_resp.insert(0, valores[4])
        self.entry_pub.delete(0, tk.END)
        self.entry_pub.insert(0, valores[5])

    def _salvar_alteracoes(self):
        """Salva alterações do registro em edição"""
        if self.dado_editando is None:
            messagebox.showwarning("Aviso", "Nenhum registro selecionado para edição.")
            return

        nome = self.entry_nome.get().strip()
        processo = self.entry_processo.get().strip()
        tipo = self.entry_tipo.get().strip()
        data = self.entry_data_prazo.get().strip()
        resp = self.entry_resp.get().strip()
        pub = self.entry_pub.get().strip()
        obs = self.entry_obs.get("1.0", tk.END).strip()
        data_notificar = self.entry_data_notificar.get().strip()

        indice = self.tree.index(self.dado_editando)
        self.registros[indice] = (nome, processo, tipo, data, data_notificar, resp, pub, obs)

        obs_display = obs[:30] + ("..." if len(obs) > 30 else "")
        self.tree.item(self.dado_editando, values=(nome, processo, tipo, data, resp, pub, obs_display))

        self.dado_editando = None
        self._limpar_campos()
        messagebox.showinfo("Sucesso", "Registro atualizado!")

    def _excluir(self):
        """Exclui registro selecionado"""
        item = self.tree.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione um registro para excluir.")
            return

        if messagebox.askyesno("Confirmação", "Deseja excluir este registro?"):
            indice = self.tree.index(item[0])
            self.registros.pop(indice)
            self.tree.delete(item)
            messagebox.showinfo("Sucesso", "Registro excluído!")

    def _limpar_campos(self):
        """Limpa todos os campos de entrada"""
        self.entry_data_prazo.delete(0, tk.END)
        self.entry_nome.delete(0, tk.END)
        self.entry_processo.delete(0, tk.END)
        self.entry_tipo.delete(0, tk.END)
        self.entry_resp.delete(0, tk.END)
        self.entry_data_notificar.delete(0, tk.END)
        self.entry_pub.delete(0, tk.END)
        self.entry_obs.delete("1.0", tk.END)

    def _gerar_excel(self):
        """Gera arquivo Excel com os registros"""
        if not self.registros:
            messagebox.showerror("Erro", "Nenhum dado para preencher.")
            return

        # Selecionar local de salvamento
        caminho_saida = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Prazos_Preenchidos.xlsx"
        )

        if not caminho_saida:
            return

        # Preencher Excel
        if self.excel_handler.preencher_prazos(self.registros, caminho_saida):
            # Não precisa registrar de novo (já salvamos ao adicionar)
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{caminho_saida}")
            self.registros.clear()
            self.tree.delete(*self.tree.get_children())
        else:
            messagebox.showerror("Erro", "Erro ao preencher planilha.")

    def _registrar_prazos_json(self):
        """Registra prazos em JSON após gerar Excel"""
        for nome, processo, tipo, data, data_notificar, resp, pub, obs in self.registros:
            prazo = {
                "cliente": nome,
                "processo": processo,
                "tipo_prazo": tipo,
                "data_fatal": data,
                "data_para_notificar": data_notificar,
                "lembrete": "FATAL",
                "notificado": False,
                "registro_em": datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            self.storage.adicionar_prazo(prazo)
