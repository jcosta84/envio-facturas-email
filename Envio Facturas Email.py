import pandas as pd
import pyodbc
import smtplib
import os
import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.ttk import Progressbar, Notebook, Frame, Label, Entry, Button, Treeview
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

relatorio_envio = []

# Janela principal
root = tk.Tk()
root.title("Gestão de Clientes e Envio de E-mails")
root.geometry("900x600")
nb = Notebook(root)
nb.pack(fill='both', expand=True)

# Variáveis ocultas (credenciais)
server_var = tk.StringVar(value="192.168.52.180,1433")
db_var = tk.StringVar(value="factura_email")
user_var = tk.StringVar(value="sa")
pwd_var = tk.StringVar(value="loucoste9850053")
gmail_var = tk.StringVar(value="cpjcosta30@gmail.com")
gmail_pwd_var = tk.StringVar(value="ilfr gubf rcfr tyro")

def conectar_sql(server, database, username, password):
    try:
        driver = "ODBC Driver 17 for SQL Server"
        connection_string = (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password}"
        )
        conn = pyodbc.connect(connection_string)
        return conn
    except Exception as e:
        messagebox.showerror("Erro", f"Erro na conexão: {e}")
        return None

def cadastrar_cliente():
    conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO clientes (cil, nome, email, arquivo_anexo) VALUES (?, ?, ?, ?)", 
                           entry_cil.get(), entry_nome.get(), entry_email_cli.get(), entry_anexo.get())
            conn.commit()
            messagebox.showinfo("Sucesso", "Cliente cadastrado!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao cadastrar: {e}")
        conn.close()

def enviar_email(destinatario, nome, cil, remetente, senha_app, caminho_anexo=None):
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = "Factura de Energia da Empresa EDEC"
    corpo = f"Olá {nome},\n\nSeu CIL é: {cil}.\n\nEis a factura do mês correspondente ao seu local\n\nAtenciosamente,\nDepartamento Gestão de Contagens EDEC SUL"
    msg.attach(MIMEText(corpo, 'plain'))

    if caminho_anexo and os.path.isfile(caminho_anexo):
        with open(caminho_anexo, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(caminho_anexo))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(remetente, senha_app)
            smtp.send_message(msg)

        conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
        if conn:
            cursor = conn.cursor()
            data_envio = datetime.datetime.now()
            cursor.execute(
                "INSERT INTO relatorio (nome, email, cil, status, mensagem, data_envio) VALUES (?, ?, ?, ?, ?, ?)",
                (nome, destinatario, cil, "Enviado", "Email enviado com sucesso", data_envio)
            )
            conn.commit()
            conn.close()
        return 'Sucesso', 'Email enviado com sucesso'

    except Exception as e:
        erro = str(e)
        conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
        if conn:
            cursor = conn.cursor()
            data_envio = datetime.datetime.now()
            cursor.execute(
                "INSERT INTO relatorio (nome, email, cil, status, mensagem, data_envio) VALUES (?, ?, ?, ?, ?, ?)",
                (nome, destinatario, cil, "Erro", erro, data_envio)
            )
            conn.commit()
            conn.close()
        return 'Erro', erro

def enviar_emails():
    relatorio_envio.clear()
    pasta = entry_pasta.get()
    conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
    if not conn: return

    cursor = conn.cursor()
    cursor.execute("SELECT cil, nome, email, arquivo_anexo FROM clientes")
    dados = cursor.fetchall()

    total = len(dados)
    barra['maximum'] = total
    barra['value'] = 0

    for i, (cil, nome, email, anexo) in enumerate(dados):
        caminho = os.path.join(pasta, anexo) if anexo else ''
        if not os.path.isfile(caminho):
            continue
        status, msg = enviar_email(email, nome, cil, gmail_var.get(), gmail_pwd_var.get(), caminho)
        relatorio_envio.append((nome, email, cil, status, msg, datetime.datetime.now()))
        barra['value'] = i + 1
        root.update()

    conn.close()
    messagebox.showinfo("Concluído", "Envio finalizado!")

def salvar_relatorio():
    conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
    if conn:
        cursor = conn.cursor()
        for r in relatorio_envio:
            cursor.execute("""
                INSERT INTO relatorio (nome, email, cil, status, mensagem, data_envio)
                VALUES (?, ?, ?, ?, ?, ?)""", r)
        conn.commit()
        conn.close()

def carregar_relatorio():
    try:
        conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
        if conn:
            cursor = conn.cursor()
            cursor.execute("SELECT nome, email, cil, status, mensagem, data_envio FROM relatorio ORDER BY data_envio DESC")
            rows = cursor.fetchall()
            for i in tree.get_children():
                tree.delete(i)
            for row in rows:
                valores = [str(x) if x is not None else '' for x in row]
                tree.insert('', 'end', values=valores)
            conn.close()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar relatório:\n{e}")

def exportar_excel():
    conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
    if conn:
        try:
            df = pd.read_sql("SELECT nome, email, cil, status, mensagem, data_envio FROM relatorio", conn)
            caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if caminho:
                df.to_excel(caminho, index=False)
                messagebox.showinfo("Sucesso", f"Relatório salvo em: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar Excel: {e}")
        finally:
            conn.close()

def exportar_pdf():
    conn = conectar_sql(server_var.get(), db_var.get(), user_var.get(), pwd_var.get())
    if conn:
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT nome, email, cil, status, mensagem, data_envio FROM relatorio")
            rows = cursor.fetchall()
            caminho = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not caminho:
                return

            c = canvas.Canvas(caminho, pagesize=A4)
            largura, altura = A4
            y = altura - 40
            c.setFont("Helvetica", 10)

            for row in rows:
                linha = f"{row[0]} | {row[1]} | CIL: {row[2]} | {row[3]} | {row[4][:40]} | {row[5].strftime('%d/%m/%Y %H:%M')}"
                c.drawString(40, y, linha)
                y -= 15
                if y < 50:
                    c.showPage()
                    c.setFont("Helvetica", 10)
                    y = altura - 40

            c.save()
            messagebox.showinfo("Sucesso", f"PDF salvo: {caminho}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar PDF: {e}")
        finally:
            conn.close()

# Aba 1: Cadastro
aba1 = Frame(nb)
nb.add(aba1, text="Cadastro")

Label(aba1, text="CIL").grid(row=0, column=0)
entry_cil = Entry(aba1)
entry_cil.grid(row=0, column=1)

Label(aba1, text="Nome").grid(row=1, column=0)
entry_nome = Entry(aba1)
entry_nome.grid(row=1, column=1)

Label(aba1, text="Email").grid(row=2, column=0)
entry_email_cli = Entry(aba1)
entry_email_cli.grid(row=2, column=1)

Label(aba1, text="Nome PDF").grid(row=3, column=0)
entry_anexo = Entry(aba1)
entry_anexo.grid(row=3, column=1)

def atualizar_nome_pdf(event=None):
    cil = entry_cil.get().strip()
    if cil:
        entry_anexo.delete(0, tk.END)
        entry_anexo.insert(0, f"{cil}.pdf")

entry_cil.bind("<KeyRelease>", atualizar_nome_pdf)

Button(aba1, text="Cadastrar", command=cadastrar_cliente).grid(row=4, column=0, columnspan=2, pady=10)

# Aba 2: Envio (oculta campos de configuração)
aba2 = Frame(nb)
nb.add(aba2, text="Envio")

Label(aba2, text="Pasta Anexos").grid(row=0, column=0)
entry_pasta = Entry(aba2)
entry_pasta.grid(row=0, column=1)
Button(aba2, text="Selecionar", command=lambda: entry_pasta.insert(0, filedialog.askdirectory())).grid(row=0, column=2)

Button(aba2, text="Enviar E-mails", command=enviar_emails).grid(row=1, column=0, columnspan=3, pady=10)
barra = Progressbar(aba2, orient='horizontal', length=300, mode='determinate')
barra.grid(row=2, column=0, columnspan=3, pady=5)

# Aba 3: Relatório
aba3 = Frame(nb)
nb.add(aba3, text="Relatório")
tree = Treeview(aba3, columns=("nome", "email", "cil", "status", "mensagem", "data_envio"), show="headings")
for col in tree['columns']:
    tree.heading(col, text=col)
tree.pack(fill='both', expand=True, padx=10, pady=10)
Button(aba3, text="Atualizar", command=carregar_relatorio).pack(pady=5)
frame_btn = tk.Frame(aba3)
frame_btn.pack()
Button(frame_btn, text="Exportar Excel", command=exportar_excel).pack(side='left', padx=10)
Button(frame_btn, text="Exportar PDF", command=exportar_pdf).pack(side='left', padx=10)

root.mainloop()
