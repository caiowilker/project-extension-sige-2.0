import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from sige_automator import SigeAutomator


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Automação SIGE - Interface Moderna")
        self.root.geometry("600x500")
        self.excel_path = ""
        self.navegador = tb.StringVar(value='chrome')
        self.usuario = tb.StringVar()
        self.senha = tb.StringVar()

        self.criar_interface()

    def criar_interface(self):
        # Título
        tb.Label(self.root, text="Automação SIGE", font=("Segoe UI", 22, "bold"), bootstyle="primary").pack(pady=15)

        # Login
        frame_login = tb.Frame(self.root, padding=10)
        frame_login.pack(fill=X, padx=30)

        tb.Label(frame_login, text="Usuário:", font=("Segoe UI", 12)).pack(anchor=W)
        tb.Entry(frame_login, textvariable=self.usuario, width=40).pack(anchor=W, pady=5)

        tb.Label(frame_login, text="Senha:", font=("Segoe UI", 12)).pack(anchor=W)
        tb.Entry(frame_login, textvariable=self.senha, width=40, show="*").pack(anchor=W, pady=5)

        # Seletor de arquivo
        frame_arquivo = tb.Frame(self.root, padding=10)
        frame_arquivo.pack(fill=X, padx=30, pady=10)

        tb.Label(frame_arquivo, text="Arquivo Excel:", font=("Segoe UI", 12)).pack(anchor=W)
        self.entrada_excel = tb.Entry(frame_arquivo, width=50)
        self.entrada_excel.pack(side=LEFT, padx=(0, 10), pady=5)

        tb.Button(frame_arquivo, text="Selecionar", bootstyle="info", command=self.selecionar_arquivo).pack(side=LEFT)

        # Seleção do navegador
        frame_nav = tb.Frame(self.root, padding=10)
        frame_nav.pack(fill=X, padx=30)

        tb.Label(frame_nav, text="Selecionar Navegador:", font=("Segoe UI", 12)).pack(anchor=W)
        tb.OptionMenu(frame_nav, self.navegador, self.navegador.get(), 'chrome', 'edge').pack(anchor=W, pady=5)

        # Botões
        botoes = tb.Frame(self.root, padding=20)
        botoes.pack()

        tb.Button(botoes, text="Executar Automação", bootstyle="success,outline", width=20,
                  command=self.executar_automacao).pack(side=LEFT, padx=10)
        tb.Button(botoes, text="Abrir Log", bootstyle="secondary,outline", width=15, command=self.abrir_log).pack(
            side=LEFT, padx=10)

        # Status
        self.status_label = tb.Label(self.root, text="", font=("Segoe UI", 11), bootstyle="info")
        self.status_label.pack(pady=20)

    def selecionar_arquivo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if file_path:
            self.excel_path = file_path
            self.entrada_excel.delete(0, tb.END)
            self.entrada_excel.insert(0, file_path)

    def executar_automacao(self):
        if not self.excel_path:
            messagebox.showerror("Erro", "Selecione um arquivo Excel primeiro.")
            return

        if not self.usuario.get() or not self.senha.get():
            messagebox.showerror("Erro", "Preencha o usuário e a senha.")
            return

        try:
            self.status_label.config(text="⏳ Executando automação...")
            self.root.update()

            automator = SigeAutomator(
                excel_path=self.excel_path,
                navegador=self.navegador.get(),
                usuario=self.usuario.get(),
                senha=self.senha.get()
            )
            automator.executar()

            self.status_label.config(text="✅ Automação concluída com sucesso!")
            abrir = messagebox.askyesno("Relatório", "Deseja abrir o relatório agora?")
            if abrir:
                os.startfile("relatorio_final.xlsx")

        except Exception as e:
            self.status_label.config(text="❌ Erro na automação.")
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

    def abrir_log(self):
        try:
            os.startfile("sige_automacao.log")
        except Exception:
            messagebox.showerror("Erro", "Log não encontrado.")


if __name__ == "__main__":
    app_style = tb.Style(theme="flatly")
    root = app_style.master
    App(root)
    root.mainloop()
