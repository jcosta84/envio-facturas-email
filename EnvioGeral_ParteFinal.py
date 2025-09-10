from urllib.parse import quote_plus
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from sqlalchemy import create_engine, text
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib
import os
import datetime
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from tkinter import ttk
import tkinter.filedialog as fd


# ========== Configurações iniciais ===========
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gestão de Clientes e Envio de E-mails")
        self.geometry("1000x700")

        # Conexão ao banco
        self.engine = self.get_engine(
            host="192.168.38.2",
            port="1433",
            database="factura_email",
            username="pcosta",
            password="loucoste9850053"
        )

        # Login do remetente
        self.remetente = "edecsul@gmail.com"
        self.senha_app = "mjmg wmua fhpk vwbv"

        # Layout
        self.menu_lateral = ctk.CTkFrame(self, width=200)
        self.menu_lateral.pack(side="left", fill="y")

        self.container = ctk.CTkFrame(self)
        self.container.pack(side="right", expand=True, fill="both")

        # Botões do menu
        botoes = [
            ("📋 Cadastro", self.abrir_cadastro),
            ("📧 CC Conhecimento", self.abrir_cc),
            ("🔍 Consulta", self.abrir_consulta),
            ("📤 Enviar E-mails", self.abrir_envio),
            ("📊 Relatório", self.abrir_relatorio)
        ]
        for texto, comando in botoes:
            btn = ctk.CTkButton(self.menu_lateral, text=texto, command=comando)
            btn.pack(pady=10, padx=10, fill="x")

        # Inicialização com tela de cadastro
        self.abrir_cadastro()

    # ========== Conexão DB ==========
    def get_engine(self, host, database, username, password, port="1433"):
        from urllib.parse import quote_plus
        try:
            password_enconde = quote_plus(password)
            connection_url = (
                f"mssql+pyodbc://{username}:{password_enconde}@{host},{port}/{database}"
                "?driver=ODBC+Driver+17+for+SQL+Server"
            )
            engine = create_engine(connection_url)
            return engine
        except Exception as e:
            messagebox.showerror("Erro de conexão", str(e))
            return None

    # ========== Navegação ==========
    def limpar_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def abrir_cadastro(self):
        self.limpar_container()
        CadastroFrame(self.container, self.engine).pack(expand=True, fill="both")

    def abrir_cc(self):
        self.limpar_container()
        CCFrame(self.container, self.engine).pack(expand=True, fill="both")

    def abrir_consulta(self):
        self.limpar_container()
        ConsultaFrame(self.container, self.engine).pack(expand=True, fill="both")

    def abrir_envio(self):
        self.limpar_container()
        EnvioFrame(self.container, self.engine, self.remetente, self.senha_app).pack(expand=True, fill="both")

    def abrir_relatorio(self):
        self.limpar_container()
        RelatorioFrame(self.container, self.engine).pack(expand=True, fill="both")

#=============== cadastro ====================
class CadastroFrame(ctk.CTkFrame):
    def __init__(self, master, engine):
        super().__init__(master)
        self.engine = engine

        ctk.CTkLabel(self, text="📋 Cadastro de Cliente", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        self.form_frame = ctk.CTkFrame(self)
        self.form_frame.pack(pady=10)

        # Campos do formulário
        self.cil_var = tk.StringVar()
        self.nome_var = tk.StringVar()
        self.email_var = tk.StringVar()

        ctk.CTkLabel(self.form_frame, text="CIL").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.cil_entry = ctk.CTkEntry(self.form_frame, textvariable=self.cil_var, width=250)
        self.cil_entry.grid(row=0, column=1, padx=10, pady=5)

        ctk.CTkLabel(self.form_frame, text="Nome").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.nome_entry = ctk.CTkEntry(self.form_frame, textvariable=self.nome_var, width=250)
        self.nome_entry.grid(row=1, column=1, padx=10, pady=5)

        ctk.CTkLabel(self.form_frame, text="Email").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.email_entry = ctk.CTkEntry(self.form_frame, textvariable=self.email_var, width=250)
        self.email_entry.grid(row=2, column=1, padx=10, pady=5)

        ctk.CTkLabel(self.form_frame, text="Nome do PDF").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.nome_pdf_label = ctk.CTkLabel(self.form_frame, text="", text_color="gray")
        self.nome_pdf_label.grid(row=3, column=1, padx=10, pady=5)

        # Botões
        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(pady=15)

        ctk.CTkButton(self.btn_frame, text="Cadastrar Cliente", command=self.cadastrar_cliente).grid(row=0, column=0, padx=10)
        ctk.CTkButton(self.btn_frame, text="Limpar", command=self.limpar_campos).grid(row=0, column=1, padx=10)

        # Atualiza nome do PDF quando CIL muda
        self.cil_var.trace_add("write", self.atualizar_nome_pdf)

    def atualizar_nome_pdf(self, *args):
        cil = self.cil_var.get()
        nome_pdf = f"{cil}.pdf" if cil else ""
        self.nome_pdf_label.configure(text=nome_pdf)

    def limpar_campos(self):
        self.cil_var.set("")
        self.nome_var.set("")
        self.email_var.set("")
        self.nome_pdf_label.configure(text="")

    def cadastrar_cliente(self):
        cil = self.cil_var.get().strip()
        nome = self.nome_var.get().strip()
        email = self.email_var.get().strip()
        nome_pdf = f"{cil}.pdf"

        if not cil or not nome or not email:
            messagebox.showwarning("Campos obrigatórios", "Por favor, preencha todos os campos.")
            return

        try:
            with self.engine.begin() as conn:
                result = conn.execute(
                    text("SELECT COUNT(*) FROM clientes WHERE cil = :cil"),
                    {"cil": cil}
                )
                count = result.scalar()

                if count > 0:
                    messagebox.showwarning("Duplicado", "Cliente com este CIL já está cadastrado.")
                else:
                    conn.execute(
                        text("INSERT INTO clientes (cil, nome, email, arquivo_anexo) VALUES (:cil, :nome, :email, :anexo)"),
                        {"cil": cil, "nome": nome, "email": email, "anexo": nome_pdf}
                    )
                    messagebox.showinfo("Sucesso", "Cliente cadastrado com sucesso!")
                    self.limpar_campos()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao cadastrar: {e}")

#========= para conhecimento==================
class CCFrame(ctk.CTkFrame):
    def __init__(self, master, engine):
        super().__init__(master)
        self.engine = engine
        self.df = pd.DataFrame()
        self.edits = {}
        self.check_vars = {}

        ctk.CTkLabel(self, text="📧 E-mails para Conhecimento (CC)", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        # Tabela de CC
        self.tabela_frame = ctk.CTkScrollableFrame(self, height=300)
        self.tabela_frame.pack(padx=20, pady=10, fill="both")

        self.carregar_dados()

        # Formulário de novo e-mail
        ctk.CTkLabel(self, text="➕ Adicionar novo e-mail CC").pack(pady=(20, 5))
        self.novo_email_var = tk.StringVar()
        ctk.CTkEntry(self, textvariable=self.novo_email_var, width=300).pack()
        ctk.CTkButton(self, text="Adicionar", command=self.adicionar_email).pack(pady=10)

        # Botões de ações
        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(pady=15)

        ctk.CTkButton(self.btn_frame, text="🗑️ Excluir Selecionados", command=self.excluir_selecionados).grid(row=0, column=0, padx=10)
        ctk.CTkButton(self.btn_frame, text="💾 Salvar Alterações", command=self.salvar_edicoes).grid(row=0, column=1, padx=10)

    def carregar_dados(self):
        for widget in self.tabela_frame.winfo_children():
            widget.destroy()

        try:
            self.df = pd.read_sql("SELECT * FROM cc_email", self.engine)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
            return

        self.check_vars = {}
        self.edits = {}

        for i, row in self.df.iterrows():
            var = tk.BooleanVar()
            self.check_vars[row["id"]] = var

            email_var = tk.StringVar(value=row["email_cc"])
            self.edits[row["id"]] = email_var

            ctk.CTkCheckBox(self.tabela_frame, text="", variable=var).grid(row=i, column=0, padx=5)
            ctk.CTkEntry(self.tabela_frame, textvariable=email_var, width=300).grid(row=i, column=1, padx=5)

    def adicionar_email(self):
        novo = self.novo_email_var.get().strip()
        if not novo:
            messagebox.showwarning("Atenção", "Informe o e-mail antes de adicionar.")
            return
        try:
            df_novo = pd.DataFrame({"email_cc": [novo]})
            df_novo.to_sql("cc_email", con=self.engine, if_exists="append", index=False)
            self.novo_email_var.set("")
            self.carregar_dados()
            messagebox.showinfo("Sucesso", "E-mail adicionado com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar e-mail: {e}")

    def excluir_selecionados(self):
        selecionados = [k for k, v in self.check_vars.items() if v.get()]
        if not selecionados:
            messagebox.showwarning("Atenção", "Nenhum e-mail selecionado.")
            return
        try:
            with self.engine.begin() as conn:
                for id_ in selecionados:
                    conn.execute(text("DELETE FROM cc_email WHERE id = :id"), {"id": id_})
            self.carregar_dados()
            messagebox.showinfo("Sucesso", "E-mails excluídos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir: {e}")

    def salvar_edicoes(self):
        try:
            with self.engine.begin() as conn:
                for id_, var in self.edits.items():
                    conn.execute(
                        text("UPDATE cc_email SET email_cc = :email WHERE id = :id"),
                        {"email": var.get().strip(), "id": id_}
                    )
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar alterações: {e}")

#========= consultar cadastro =================
class ConsultaFrame(ctk.CTkFrame):
    def __init__(self, master, engine):
        super().__init__(master)
        self.engine = engine
        self.df_original = pd.DataFrame()
        self.edits = {}
        self.check_vars = {}

        ctk.CTkLabel(self, text="🔍 Consultar Cadastro de Clientes", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        self.filtro_frame = ctk.CTkFrame(self)
        self.filtro_frame.pack(pady=10)

        ctk.CTkLabel(self.filtro_frame, text="Filtrar por CIL:").grid(row=0, column=0, padx=10, pady=5)

        self.cil_filtro = ctk.CTkComboBox(self.filtro_frame, values=["Todos"], command=self.filtrar)
        self.cil_filtro.grid(row=0, column=1, padx=10, pady=5)

        self.tabela_frame = ctk.CTkScrollableFrame(self, height=350)
        self.tabela_frame.pack(pady=10, fill="both", expand=True)

        # Botões
        btns = ctk.CTkFrame(self)
        btns.pack(pady=10)
        ctk.CTkButton(btns, text="🗑️ Excluir Selecionados", command=self.excluir_selecionados).grid(row=0, column=0, padx=10)
        ctk.CTkButton(btns, text="💾 Salvar Alterações", command=self.salvar_alteracoes).grid(row=0, column=1, padx=10)

        self.carregar_dados()

    def carregar_dados(self):
        for widget in self.tabela_frame.winfo_children():
            widget.destroy()

        try:
            self.df_original = pd.read_sql("SELECT cil, nome, email FROM clientes", self.engine)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {e}")
            return

        self.lista_cil = self.df_original["cil"].dropna().astype(str).tolist()
        self.cil_filtro.configure(values=["Todos"] + self.lista_cil)
        self.cil_filtro.set("Todos")

        self.exibir_tabela(self.df_original)

    def filtrar(self, valor):
        if valor == "Todos":
            self.exibir_tabela(self.df_original)
        else:
            filtrado = self.df_original[self.df_original["cil"] == valor]
            self.exibir_tabela(filtrado)

    def exibir_tabela(self, df):
        for widget in self.tabela_frame.winfo_children():
            widget.destroy()

        self.check_vars = {}
        self.edits = {}

        for i, row in df.iterrows():
            var = tk.BooleanVar()
            self.check_vars[row["cil"]] = var

            email_var = tk.StringVar(value=row["email"])
            self.edits[row["cil"]] = email_var

            ctk.CTkCheckBox(self.tabela_frame, text="", variable=var).grid(row=i, column=0, padx=5)
            ctk.CTkLabel(self.tabela_frame, text=row["cil"], width=100).grid(row=i, column=1, padx=5)
            ctk.CTkLabel(self.tabela_frame, text=row["nome"], width=200).grid(row=i, column=2, padx=5)
            ctk.CTkEntry(self.tabela_frame, textvariable=email_var, width=250).grid(row=i, column=3, padx=5)

    def excluir_selecionados(self):
        selecionados = [cil for cil, var in self.check_vars.items() if var.get()]
        if not selecionados:
            messagebox.showwarning("Atenção", "Nenhum cliente selecionado.")
            return
        try:
            with self.engine.begin() as conn:
                for cil in selecionados:
                    conn.execute(text("DELETE FROM clientes WHERE cil = :cil"), {"cil": cil})
            self.carregar_dados()
            messagebox.showinfo("Sucesso", "Clientes excluídos com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao excluir: {e}")

    def salvar_alteracoes(self):
        try:
            with self.engine.begin() as conn:
                for cil, email_var in self.edits.items():
                    conn.execute(
                        text("UPDATE clientes SET email = :email WHERE cil = :cil"),
                        {"email": email_var.get().strip(), "cil": cil}
                    )
            messagebox.showinfo("Sucesso", "Alterações salvas com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar alterações: {e}")

#============ envio email =====================
class EnvioFrame(ctk.CTkFrame):
    def __init__(self, master, engine, remetente, senha_app):
        super().__init__(master)
        self.engine = engine
        self.remetente = remetente
        self.senha_app = senha_app
        self.pdf_dict = {}

        ctk.CTkLabel(self, text="📤 Envio de E-mails com Anexo PDF", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        # Botão para selecionar múltiplos arquivos PDF
        ctk.CTkButton(self, text="Selecionar Directotio", command=self.selecionar_pdfs).pack(pady=10)
        self.arquivos_label = ctk.CTkLabel(self, text="")
        self.arquivos_label.pack(pady=5)

        # Botão para iniciar envio
        ctk.CTkButton(self, text="📨 Enviar E-mails", command=self.enviar_em_lote).pack(pady=15)

        self.status_box = tk.Text(self, height=15, width=100)
        self.status_box.pack(pady=10)

    def selecionar_pdfs(self):
        
        """Selecionar diretório contendo arquivos PDF"""
        self.pdfs = filedialog.askdirectory(title="Selecionar diretório")
        if self.pdfs:
            print(f"Diretório selecionado: {self.pdfs}")
            # Lista os PDFs do diretório e atualiza a label e dicionário
            self.pdf_dict.clear()
            nomes = []
            for nome_arquivo in os.listdir(self.pdfs):
                if nome_arquivo.lower().endswith('.pdf'):
                    caminho_completo = os.path.join(self.pdfs, nome_arquivo)
                    self.pdf_dict[nome_arquivo] = caminho_completo
                    nomes.append(nome_arquivo)
            self.arquivos_label.configure(text=f"{len(nomes)} arquivo(s) PDF encontrado(s)")



    def enviar_em_lote(self):
        if not self.pdf_dict:
            messagebox.showwarning("Atenção", "Nenhum arquivo PDF selecionado.")
            return

        try:
            df_clientes = pd.read_sql("SELECT cil, nome, email, arquivo_anexo FROM clientes", self.engine)
            df_cc = pd.read_sql("SELECT email_cc FROM cc_email", self.engine)
            emails_cc = df_cc["email_cc"].dropna().tolist()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")
            return

        relatorio_envio = []
        self.status_box.delete(1.0, tk.END)

        for idx, row in df_clientes.iterrows():
            cil, nome, email, anexo = row["cil"], row["nome"], row["email"], row["arquivo_anexo"]
            caminho_pdf = self.pdf_dict.get(anexo)

            if not caminho_pdf or not os.path.isfile(caminho_pdf):
                self.status_box.insert(tk.END, f"[IGNORADO] {cil} - Arquivo '{anexo}' não encontrado.\n")
                continue

            status, mensagem = self.enviar_email(email, nome, cil, caminho_pdf, emails_cc)
            relatorio_envio.append((nome, email, cil, status, mensagem, datetime.datetime.now()))

            self.status_box.insert(tk.END, f"[{status}] {nome} - {mensagem}\n")
            self.status_box.see(tk.END)
            self.update()

        # Gravar no banco de dados
        try:
            with self.engine.begin() as conn:
                for r in relatorio_envio:
                    conn.execute(text("""
                        INSERT INTO relatorio (nome, email, cil, status, mensagem, data_envio)
                        VALUES (:nome, :email, :cil, :status, :mensagem, :data_envio)
                    """), {
                        "nome": r[0], "email": r[1], "cil": r[2],
                        "status": r[3], "mensagem": r[4], "data_envio": r[5]
                    })
            messagebox.showinfo("Concluído", "Todos os e-mails foram processados.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar relatório: {e}")

    def obter_corpo_padrao(self):
        with self.engine.connect() as conn:
            result = conn.execute(text("SELECT conteudo FROM corpo_email WHERE id = 1"))
            row = result.fetchone()
            return row[0] if row else ""
    
    def enviar_email(self, destinatario, nome, cil, caminho_anexo, cc_list=None):
        cc_list = cc_list or []

        #buscar coprp de texto
        corpo_padrao = self.obter_corpo_padrao()

        msg = MIMEMultipart()
        msg['From'] = self.remetente
        msg['To'] = destinatario
        msg['Subject'] = "Factura de Energia da Empresa EDEC"
        if cc_list:
            msg['Cc'] = ", ".join(cc_list)

        corpo = f"""
        Olá {nome},
       
        📄 CIL: {cil}

        {corpo_padrao}
       
        """
        msg.attach(MIMEText(corpo, 'plain'))

        with open(caminho_anexo, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(caminho_anexo))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
        msg.attach(part)

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(self.remetente, self.senha_app)
                smtp.send_message(msg)
            return "Sucesso", "Email enviado com sucesso."
        except Exception as e:
            return "Erro", str(e)

#========= relatorio ==========================
class RelatorioFrame(ctk.CTkFrame):
    def __init__(self, master, engine):
        super().__init__(master)
        self.engine = engine
        self.df = pd.DataFrame()

        ctk.CTkLabel(self, text="📊 Relatório de Envios", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        self.filtro_frame = ctk.CTkFrame(self)
        self.filtro_frame.pack(pady=10)

        self.inicio_var = tk.StringVar()
        self.fim_var = tk.StringVar()

        ctk.CTkLabel(self.filtro_frame, text="Data Início (AAAA-MM-DD):").grid(row=0, column=0, padx=5)
        ctk.CTkEntry(self.filtro_frame, textvariable=self.inicio_var, width=120).grid(row=0, column=1, padx=5)

        ctk.CTkLabel(self.filtro_frame, text="Data Fim (AAAA-MM-DD):").grid(row=0, column=2, padx=5)
        ctk.CTkEntry(self.filtro_frame, textvariable=self.fim_var, width=120).grid(row=0, column=3, padx=5)

        ctk.CTkButton(self.filtro_frame, text="🔍 Filtrar", command=self.filtrar).grid(row=0, column=4, padx=10)

        self.tabela = tk.Text(self, height=15, width=120)
        self.tabela.pack(pady=10)

        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(pady=10)

        ctk.CTkButton(self.btn_frame, text="📥 Exportar Excel", command=self.exportar_excel).grid(row=0, column=0, padx=10)
        ctk.CTkButton(self.btn_frame, text="📄 Exportar PDF", command=self.exportar_pdf).grid(row=0, column=1, padx=10)

        self.carregar_dados()

    def carregar_dados(self):
        try:
            self.df = pd.read_sql("SELECT nome, email, cil, status, mensagem, data_envio FROM relatorio ORDER BY data_envio DESC", self.engine)
            self.df["data_envio"] = pd.to_datetime(self.df["data_envio"], errors="coerce")
            self.exibir()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar relatório: {e}")

    def filtrar(self):
        inicio = self.inicio_var.get().strip()
        fim = self.fim_var.get().strip()

        try:
            df_filtrado = self.df.copy()
            if inicio:
                inicio_dt = pd.to_datetime(inicio).normalize()
                df_filtrado = df_filtrado[df_filtrado["data_envio"] >= inicio_dt]

            if fim:
                fim_dt = pd.to_datetime(fim).normalize() + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                df_filtrado = df_filtrado[df_filtrado["data_envio"] <= fim_dt]
            self.exibir(df_filtrado)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao aplicar filtro: {e}")

    def exibir(self, df=None):
        df = df if df is not None else self.df
        self.tabela.delete(1.0, tk.END)
        for _, row in df.iterrows():
            linha = f"{row['data_envio'].strftime('%Y-%m-%d %H:%M')} | {row['nome']} | {row['email']} | CIL: {row['cil']} | {row['status']}\n"
            self.tabela.insert(tk.END, linha)

    def exportar_excel(self):
        if self.df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if not file_path:
                return
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                self.df.to_excel(writer, index=False, sheet_name="Relatório")
            messagebox.showinfo("Sucesso", f"Arquivo salvo: {file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar: {e}")

    def exportar_pdf(self):
        if self.df.empty:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return

        try:
            file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if not file_path:
                return

            buffer = BytesIO()
            c = canvas.Canvas(buffer, pagesize=A4)
            largura, altura = A4
            y = altura - 40
            c.setFont("Helvetica", 10)

            for _, row in self.df.iterrows():
                linha = f"{row['data_envio'].strftime('%d/%m/%Y %H:%M')} | {row['nome']} | {row['email']} | CIL: {row['cil']} | {row['status']}"
                c.drawString(40, y, linha)
                y -= 15
                if y < 50:
                    c.showPage()
                    c.setFont("Helvetica", 10)
                    y = altura - 40
            c.save()

            with open(file_path, "wb") as f:
                f.write(buffer.getvalue())

            messagebox.showinfo("Sucesso", f"PDF salvo: {file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar PDF: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
