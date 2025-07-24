import tempfile
import streamlit as st
import pandas as pd
from sqlalchemy import create_engine, text
import smtplib
import os
import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from io import BytesIO
from streamlit_option_menu import option_menu



relatorio_envio = []

# ========== Funções ==========

def get_engine(server, database, username, password):
    try:
        connection_string = f"mssql+pyodbc://{username}:{password}@{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        engine = create_engine(connection_string)
        return engine
    except Exception as e:
        st.error(f"Erro na conexão com o banco: {e}")
        return None

def enviar_email(destinatario, nome, cil, remetente, senha_app, caminho_anexo=None):
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = "Factura de Energia da Empresa EDEC"

    if lista_cc:
        msg['Cc'] = ', '.join(lista_cc)
        destinatarios_finais = [destinatario] + lista_cc
    else:
        destinatarios_finais = [destinatario]
    
    corpo = f"""
    Olá {nome},

    Informamos que sua fatura de energia já está disponível.

    📄 CIL: {cil}

    Anexamos o documento correspondente ao mês atual para que possa efetuar o pagamento.

    ⚠️ *Lembramos que o não pagamento poderá resultar na interrupção do fornecimento.*

    Caso já tenha efetuado o pagamento, por favor, desconsidere esta mensagem.

    Atenciosamente,  
    **Direção Comercial - EDEC SUL**
    """
    msg.attach(MIMEText(corpo, 'plain'))

    if caminho_anexo and os.path.isfile(caminho_anexo):
        with open(caminho_anexo, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(caminho_anexo))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho_anexo)}"'
        msg.attach(part)

    try:
        with smtplib.SMTP('smtp.office365.com', 587) as smtp:
            smtp.starttls()
            smtp.login(remetente, senha_app)
            smtp.sendmail(remetente, destinatarios_finais, msg.as_string())

        return 'Sucesso', 'Email enviado com sucesso'
    except Exception as e:
        return 'Erro', str(e)

# ========== Interface ==========

st.set_page_config("Gestão de Clientes", layout="wide")
st.title("📧 Gestão de Clientes e Envio de E-mails")
with st.sidebar:
    aba = option_menu(
        menu_title="📋 Menu",
        options=["Cadastro", "Para Conhecimento","Consultar Cadastro","Envio de E-mails", "Relatório"],
        icons=["person-plus", "envelope", "bar-chart"],
        menu_icon="cast",
        default_index=0,
    )

# Configuração
#st.sidebar.expander("⚙️ Configuração do sistema")
server_var ="192.168.52.180,1433"
db_var = "factura_email"
user_var = "sa"
pwd_var = "loucoste9850053"
gmail_var = "cpjcosta30@gmail.com"
gmail_pwd_var = "ilfr gubf rcfr tyro"

engine = get_engine(server_var, db_var, user_var, pwd_var)


# ========== Cadastro ==========
if "limpar_form" not in st.session_state:
    st.session_state["limpar_form"] = False

if aba == "Cadastro":
    st.header("📋 Cadastro de Cliente")

    if st.session_state["limpar_form"]:
        # Limpar valores
        st.session_state["cil"] = ""
        st.session_state["nome"] = ""
        st.session_state["email"] = ""
        st.session_state["limpar_form"] = False

    with st.form("form_cadastro"):
        cil = st.text_input("CIL", value=st.session_state.get("cil", ""), key="cil")
        nome = st.text_input("Nome", value=st.session_state.get("nome", ""), key="nome")
        email = st.text_input("Email", value=st.session_state.get("email", ""), key="email")

        nome_pdf = f"{cil}.pdf" if cil else ""
        st.text_input("Nome PDF", value=nome_pdf, disabled=True)

        cadastrar = st.form_submit_button("Limpar")

        if cadastrar and engine:
            if not cil.strip() or not nome.strip() or not email.strip():
                st.error("Por favor, preencha todos os campos.")
            else:
                try:
                    with engine.begin() as conn:
                        result = conn.execute(
                            text("SELECT COUNT(*) FROM clientes WHERE cil = :cil"),
                            {"cil": cil}
                        )
                        count = result.scalar()
                        if count == 0:
                            conn.execute(
                                text("INSERT INTO clientes (cil, nome, email, arquivo_anexo) VALUES (:cil, :nome, :email, :anexo)"),
                                {"cil": cil, "nome": nome, "email": email, "anexo": nome_pdf}
                            )
                            st.success("Cliente cadastrado com sucesso!")

                            # Aciona a flag para limpar na próxima execução
                            st.session_state["limpar_form"] = True
                           # st.rerun()

                        else:
                            st.warning("Cliente com este CIL já está cadastrado.")
                except Exception as e:
                    st.error(f"Erro ao cadastrar: {e}")

#=========== Conhecimento ============
if aba == "Para Conhecimento":
    st.header("📋 E-mails para Conhecimento (CC)")

    # 🔹 Carrega os dados
    query = "SELECT * FROM cc_email"
    cc = pd.read_sql(query, engine)

    if cc.empty:
        st.info("Nenhum e-mail CC cadastrado.")
    else:
        cc = cc.rename(columns={"id": "ID", "email_cc": "E-mail CC"})

        # Adiciona coluna para seleção
        cc["Selecionar"] = False

        # 🔸 Edição dos dados
        cc_editado = st.data_editor(
            cc,
            column_config={
                "E-mail CC": st.column_config.TextColumn("E-mail CC"),
                "Selecionar": st.column_config.CheckboxColumn("Selecionar para exclusão"),
            },
            use_container_width=True,
            disabled=["ID"]
        )

        col1, col2 = st.columns(2)

        # 🔸 Botão para excluir selecionados
        with col1:
            if st.button("🗑️ Excluir selecionados"):
                selecionados = cc_editado[cc_editado["Selecionar"] == True]
                if not selecionados.empty:
                    try:
                        with engine.begin() as conn:
                            for _, row in selecionados.iterrows():
                                conn.execute(text("DELETE FROM cc_email WHERE id = :id"), {"id": row["ID"]})
                        st.success(f"{len(selecionados)} e-mail(s) excluído(s) com sucesso.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao excluir: {e}")
                else:
                    st.warning("Nenhum e-mail selecionado para exclusão.")

        # 🔸 Botão para atualizar registros editados
        with col2:
            if st.button("💾 Salvar alterações"):
                try:
                    with engine.begin() as conn:
                        for _, row in cc_editado.iterrows():
                            conn.execute(
                                text("UPDATE cc_email SET email_cc = :email WHERE id = :id"),
                                {"email": row["E-mail CC"], "id": row["ID"]}
                            )
                    st.success("Alterações salvas com sucesso.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao atualizar: {e}")

    st.markdown("---")

    # 🔹 Formulário para novo e-mail CC
    st.subheader("➕ Adicionar novo e-mail CC")
    with st.form("form_cadastro"):
        novo_email = st.text_input("Novo e-mail CC", key="novo_email_cc")
        submit = st.form_submit_button("Submeter")

        if submit:
            try:
                novo = pd.DataFrame({"email_cc": [novo_email]})
                novo.to_sql("cc_email", con=engine, if_exists="append", index=False)
                st.success("E-mail CC adicionado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao inserir dados: {e}")
   
# ========== Consultar ===============
if aba == "Consultar Cadastro":
    st.header("📋 Consultar Cliente")
    query = "SELECT * FROM clientes"
    cliente = pd.read_sql(query, engine)

    cliente = cliente.rename(columns={
        "cil": "CIL",
        "nome": "Nome do Cliente",
        "email": "E-mail"
    })

    colunas_necessarias = ["CIL", "Nome do Cliente", "E-mail"]

    with st.sidebar:
        cil_opcoes = cliente["CIL"].unique()
        sel_cil = st.multiselect(
            "Selecionar CIL:",
            options=cil_opcoes,
            default=cil_opcoes
        )

    if sel_cil:
        geral_selection = cliente[cliente["CIL"].isin(sel_cil)]
    else:
        geral_selection = cliente

    df_para_exibir = geral_selection[colunas_necessarias].copy()

    # Adiciona coluna checkbox para seleção de exclusão
    if "Selecionar" not in df_para_exibir.columns:
        df_para_exibir["Selecionar"] = False

    geral_edi = st.data_editor(
        df_para_exibir,
        column_config={
            "E-mail": st.column_config.TextColumn("E-mail"),
            "Selecionar": st.column_config.CheckboxColumn("Selecionar para exclusão"),
        },
        disabled=["CIL", "Nome do Cliente"],
        use_container_width=True
    )

    if st.button("Excluir selecionados"):
        selecionados = geral_edi[geral_edi["Selecionar"] == True]
        if not selecionados.empty:
            try:
                with engine.begin() as conn:
                    for _, row in selecionados.iterrows():
                        conn.execute(
                            text("DELETE FROM clientes WHERE cil = :cil"),
                            {"cil": row["CIL"]}
                        )
                st.success(f"{len(selecionados)} cliente(s) excluído(s) com sucesso!")
                #st.experimental_rerun()
            except Exception as e:
                st.error(f"Erro ao excluir: {e}")
        else:
            st.warning("Nenhum cliente selecionado para exclusão.")

    if st.button("Salvar alterações"):
        try:
            with engine.begin() as conn:
                for _, row in geral_edi.iterrows():
                    conn.execute(
                        text("UPDATE clientes SET email = :email WHERE cil = :cil"),
                        {"email": row["E-mail"], "cil": row["CIL"]}
                    )
            st.success("Alterações salvas com sucesso!")
        except Exception as e:
            st.error(f"Erro ao salvar alterações: {e}")

# ========== Envio ===================
elif aba == "Envio de E-mails":
    st.header("📤 Envio de E-mails")

    pdfs = st.file_uploader("Selecione os arquivos PDF", type="pdf", accept_multiple_files=True)

    # Busca todos os e-mails para conhecimento (CC global)
    try:
        df_cc = pd.read_sql("SELECT email_cc FROM cc_email", engine)
        emails_cc_globais = df_cc["email_cc"].dropna().tolist()
    except Exception as e:
        st.error(f"Erro ao carregar e-mails CC: {e}")
        emails_cc_globais = []

    if st.button("Enviar E-mails") and engine:
        if not pdfs:
            st.error("Por favor, selecione pelo menos um arquivo PDF.")
        else:
            try:
                dados = pd.read_sql("SELECT cil, nome, email, arquivo_anexo FROM clientes", engine)
            except Exception as e:
                st.error(f"Erro ao carregar clientes: {e}")
                st.stop()

            total = len(dados)
            progresso = st.progress(0)
            relatorio_envio.clear()

            arquivos_pdf = {pdf.name: pdf for pdf in pdfs}

            for i, row in dados.iterrows():
                cil, nome, email, anexo = row["cil"], row["nome"], row["email"], row["arquivo_anexo"]
                if anexo and anexo in arquivos_pdf:
                    pdf_file = arquivos_pdf[anexo]

                    with tempfile.TemporaryDirectory() as temp_dir:
                        caminho_temp = os.path.join(temp_dir, f"{cil}.pdf")
                        with open(caminho_temp, "wb") as f:
                            f.write(pdf_file.getbuffer())

                        # Enviar e-mail com CC global
                        status, msg = enviar_email(email, nome, cil, gmail_var, gmail_pwd_var, caminho_temp, emails_cc_globais)

                    relatorio_envio.append((nome, email, cil, status, msg, datetime.datetime.now()))
                else:
                    continue

                progresso.progress((i + 1) / total)

            try:
                with engine.begin() as conn:
                    for r in relatorio_envio:
                        conn.execute(text("""
                            INSERT INTO relatorio (nome, email, cil, status, mensagem, data_envio)
                            VALUES (:nome, :email, :cil, :status, :mensagem, :data_envio)
                        """), {
                            "nome": r[0], "email": r[1], "cil": r[2],
                            "status": r[3], "mensagem": r[4], "data_envio": r[5]
                        })
                st.success("📨 Envio concluído com sucesso!")
            except Exception as e:
                st.error(f"Erro ao salvar relatório: {e}")

# ========== Aba: Relatório ==========
elif aba == "Relatório" and engine:
    st.header("📊 Relatório de Envios")
    try:
        df = pd.read_sql("SELECT nome, email, cil, status, mensagem, data_envio FROM relatorio ORDER BY data_envio DESC", engine)
        
        #converter data
        df['data_envio'] = pd.to_datetime(df['data_envio'], errors='coerce')  # converter para datetime
        
        if pd.api.types.is_datetime64_any_dtype(df['data_envio']):

            col1, col2 = st.columns(2)

            with col1:
                data_inicio = st.date_input("Data Início", value=df["data_envio"].min().date() if not df.empty else datetime.date.today())
            with col2:
                data_fim = st.date_input("Data Fim", value=df["data_envio"].max().date() if not df.empty else datetime.date.today())

            if data_inicio > data_fim:
                st.warning("A data de inicio não pode ser inferior a data de fim")
            else:
                df = df[(df["data_envio"].dt.date >= data_inicio) & (df["data_envio"].dt.date <= data_fim)]

        st.dataframe(df, hide_index=True)

        col1, col2 = st.columns(2)
        with col1:
            #função download excel
            buffer_excel = BytesIO()
            with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Relatório')
            buffer_excel.seek(0)
            st.download_button("📥 Exportar Excel", buffer_excel, file_name="relatorio_envio.xlsx")

        with col2:
            #função download pdf
            buffer_pdf = BytesIO()
            c = canvas.Canvas(buffer_pdf, pagesize=A4)
            largura, altura = A4
            y = altura - 40
            c.setFont("Helvetica", 10)

            for _, row in df.iterrows():
                linha = f"{row['nome']} | {row['email']} | CIL: {row['cil']} | {row['status']} | {str(row['mensagem'])[:40]} | {row['data_envio'].strftime('%d/%m/%Y %H:%M')}"
                c.drawString(40, y, linha)
                y -= 15
                if y < 50:
                    c.showPage()
                    c.setFont("Helvetica", 10)
                    y = altura - 40
            c.save()
            st.download_button("📄 Exportar PDF", buffer_pdf.getvalue(), file_name="relatorio_envio.pdf")

    except Exception as e:
        st.error(f"Erro ao carregar relatório: {e}")

