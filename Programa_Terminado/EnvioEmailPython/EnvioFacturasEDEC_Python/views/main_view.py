import re
import threading
from datetime import datetime
from pathlib import Path
from tkinter import TclError, filedialog, messagebox, ttk
from typing import Dict, List, Optional

import customtkinter as ctk

from database.repositories import (
    AuditoriaRepository,
    CcRepository,
    ClienteRepository,
    ConfiguracaoRepository,
    RelatorioRepository,
    UsuarioRepository,
)
from models.entities import Cliente, RelatorioEnvio
from services.email_service import EmailService
from services.export_service import exportar_excel, exportar_pdf

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class MainView(ctk.CTkToplevel):
    BG = "#181A1D"
    SIDEBAR = "#111315"
    PANEL = "#24272A"
    PANEL_2 = "#2D3034"
    BORDER = "#3A3E43"
    ACTIVE = "#4B4F54"
    BLUE = "#4C8DFF"

    def __init__(
        self,
        config,
        db,
        usuario,
        master=None,
        ao_terminar_sessao=None,
        ao_sair=None,
    ):
        super().__init__(master)

        self.config_service = config
        self.db = db
        self.usuario = usuario
        self.ao_terminar_sessao = ao_terminar_sessao
        self.ao_sair = ao_sair
        self._fechando = False
        self._relogio_after_id = None
        self.clientes = ClienteRepository(db)
        self.cc = CcRepository(db)
        self.conf = ConfiguracaoRepository(db)
        self.relatorios = RelatorioRepository(db)
        self.usuarios = UsuarioRepository(db)
        self.aud = AuditoriaRepository(db)
        self.email = EmailService(config)
        self.botao_ativo = None
        self.botoes_menu: Dict[str, ctk.CTkButton] = {}

        self.title("Gestão e Envio de E-mails")
        self.geometry("1280x760")
        self.minsize(1120, 700)
        self.configure(fg_color=self.BG)
        self.protocol("WM_DELETE_WINDOW", self.sair)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

        self._configurar_treeview()
        self._criar_estrutura()
        self._criar_menu()
        self._criar_status_bar()
        self.abrir_pagina("Início", self.pagina_inicio)

    def _configurar_treeview(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure(
            "Dark.Treeview",
            background="#202326",
            foreground="#FFFFFF",
            fieldbackground="#202326",
            borderwidth=0,
            rowheight=31,
            font=("Segoe UI", 10),
        )
        style.configure(
            "Dark.Treeview.Heading",
            background="#2B2E32",
            foreground="#FFFFFF",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        style.map(
            "Dark.Treeview",
            background=[("selected", self.BLUE)],
            foreground=[("selected", "#FFFFFF")],
        )
        style.map("Dark.Treeview.Heading", background=[("active", "#363A3F")])

    def _criar_estrutura(self):
        self.sidebar = ctk.CTkFrame(self, width=245, corner_radius=0, fg_color=self.SIDEBAR)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        self.content = ctk.CTkFrame(self, corner_radius=0, fg_color=self.BG)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        self.status_bar = ctk.CTkFrame(self, height=36, corner_radius=0, fg_color="#202225")
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew")
        self.status_bar.grid_columnconfigure(1, weight=1)

    def _criar_menu(self):
        ctk.CTkLabel(
            self.sidebar,
            text=self.usuario.username,
            font=ctk.CTkFont(size=18, weight="bold"),
        ).pack(pady=(28, 0))
        ctk.CTkLabel(
            self.sidebar,
            text=self.usuario.nivel.upper(),
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(0, 24))

        itens = [
            ("Início", "⌂", self.pagina_inicio),
            ("Clientes", "♟", self.pagina_clientes),
            ("Consulta", "⌕", self.pagina_consulta),
            ("Enviar E-mails", "✉", self.pagina_envio),
            ("Corpo do E-mail", "⚙", self.pagina_corpo),
            ("Gerir CC", "♟", self.pagina_cc),
            ("Relatórios", "≡", self.pagina_relatorios),
        ]
        if self.usuario.is_admin:
            itens.append(("Utilizadores", "▢", self.pagina_usuarios))
        if self.usuario.is_gerente:
            itens.append(("Definições", "⚙", self.pagina_definicoes))

        for texto, icone, comando in itens:
            botao = ctk.CTkButton(
                self.sidebar,
                text="{}   {}".format(icone, texto),
                anchor="w",
                height=38,
                corner_radius=7,
                fg_color="#34373B",
                hover_color="#45494E",
                font=ctk.CTkFont(size=13, weight="bold"),
                command=lambda n=texto, c=comando: self.abrir_pagina(n, c),
            )
            botao.pack(fill="x", padx=18, pady=5)
            self.botoes_menu[texto] = botao

        ctk.CTkButton(
            self.sidebar,
            text="↪   Sair",
            anchor="w",
            height=38,
            fg_color="#34373B",
            hover_color="#45494E",
            command=self.sair,
        ).pack(side="bottom", fill="x", padx=18, pady=(0, 18))
        ctk.CTkButton(
            self.sidebar,
            text="⏻   Terminar Sessão",
            anchor="w",
            height=38,
            fg_color="#34373B",
            hover_color="#45494E",
            command=self.terminar_sessao,
        ).pack(side="bottom", fill="x", padx=18, pady=5)
        ctk.CTkFrame(self.sidebar, height=2, fg_color="#34373B").pack(
            side="bottom", fill="x", padx=18, pady=(0, 12)
        )

    def _criar_status_bar(self):
        self.label_ligacao = ctk.CTkLabel(
            self.status_bar,
            text="Base de dados: ligada",
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family="Segoe UI",
                size=12,
                weight="bold"
            ),
        )

        self.label_ligacao.grid(
            row=0,
            column=0,
            padx=18,
            pady=8,
            sticky="w"
        )

        self.label_data_hora = ctk.CTkLabel(
            self.status_bar,
            text="",
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family="Segoe UI",
                size=12
            ),
        )

        self.label_data_hora.grid(
            row=0,
            column=2,
            padx=18,
            pady=8,
            sticky="e"
        )

        # Guarda o identificador do after para permitir o cancelamento.
        self._relogio_after_id = self.after(200, self._atualizar_relogio)

    def _atualizar_relogio(self):
        if self._fechando:
            return

        try:
            if not self.winfo_exists():
                return
            if not hasattr(self, "label_data_hora"):
                return
            if not self.label_data_hora.winfo_exists():
                return

            self.label_data_hora.configure(
                text=datetime.now().strftime("%d/%m/%Y  |  %H:%M:%S")
            )
            self._relogio_after_id = self.after(1000, self._atualizar_relogio)

        except (TclError, RuntimeError):
            self._relogio_after_id = None
        except Exception as erro:
            self._relogio_after_id = None
            print("Erro ao atualizar data e hora:", erro)

    def cancelar_relogio(self):
        if self._relogio_after_id is None:
            return
        try:
            self.after_cancel(self._relogio_after_id)
        except (TclError, RuntimeError):
            pass
        self._relogio_after_id = None

    def limpar(self):
        for widget in self.content.winfo_children():
            widget.destroy()

    def abrir_pagina(self, nome, comando):
        if self.botao_ativo is not None:
            self.botao_ativo.configure(fg_color="#34373B", border_width=0)
        botao = self.botoes_menu.get(nome)
        if botao is not None:
            botao.configure(fg_color=self.ACTIVE, border_width=3, border_color=self.BLUE)
            self.botao_ativo = botao
        comando()

    def pagina(self, titulo, scroll=True):
        self.limpar()
        if scroll:
            frame = ctk.CTkScrollableFrame(self.content, fg_color=self.BG)
        else:
            frame = ctk.CTkFrame(self.content, fg_color=self.BG)
        frame.grid(row=0, column=0, sticky="nsew", padx=26, pady=24)
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text=titulo, font=ctk.CTkFont(size=29, weight="bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 18)
        )
        return frame

    def terminar_sessao(self):
        """Fecha a sessão atual e regressa à janela de login."""
        if self._fechando:
            return

        self._fechando = True
        self.cancelar_relogio()

        try:
            self.destroy()
        except (TclError, RuntimeError):
            pass

        if callable(self.ao_terminar_sessao):
            self.ao_terminar_sessao()

    def sair(self):
        """Fecha completamente a aplicação."""
        if self._fechando:
            return

        self._fechando = True
        self.cancelar_relogio()

        if callable(self.ao_sair):
            self.ao_sair()
            return

        try:
            self.destroy()
        except (TclError, RuntimeError):
            pass

    def _criar_tree(self, parent, colunas, widths=None, height=13):
        container = ctk.CTkFrame(parent, fg_color=self.PANEL, corner_radius=10)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
        tree = ttk.Treeview(
            container,
            columns=colunas,
            show="headings",
            height=height,
            style="Dark.Treeview",
        )
        scroll_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        for index, coluna in enumerate(colunas):
            tree.heading(coluna, text=coluna)
            largura = widths[index] if widths and index < len(widths) else 140
            tree.column(coluna, width=largura, minwidth=80, anchor="center")
        tree.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=(8, 0))
        scroll_y.grid(row=0, column=1, sticky="ns", pady=(8, 0))
        scroll_x.grid(row=1, column=0, sticky="ew", padx=(8, 0), pady=(0, 8))
        return container, tree

    def _card(self, parent, column, icon, number, title, subtitle):
        card = ctk.CTkFrame(
            parent, height=155, corner_radius=12, fg_color="#292C2F",
            border_width=1, border_color="#383C40"
        )
        card.grid(row=0, column=column, sticky="nsew", padx=8, pady=5)
        card.grid_propagate(False)
        ctk.CTkLabel(card, text=icon, font=ctk.CTkFont(size=24)).pack(anchor="w", padx=18, pady=(14, 3))
        ctk.CTkLabel(card, text=str(number), font=ctk.CTkFont(size=30, weight="bold")).pack(anchor="w", padx=18)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=18)
        ctk.CTkLabel(card, text=subtitle, font=ctk.CTkFont(size=12)).pack(anchor="w", padx=18)

    def pagina_inicio(self):
        frame = self.pagina("Início")
        ctk.CTkLabel(
            frame,
            text="Bem-vindo(a) ao sistema de gestão de clientes e envio de e-mails.",
            font=ctk.CTkFont(size=15),
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))
        try:
            clientes = list(self.clientes.listar())
            relatorios = list(self.relatorios.listar())
            cc = list(self.cc.listar())
        except Exception:
            clientes, relatorios, cc = [], [], []

        cards = ctk.CTkFrame(frame, fg_color="transparent")
        cards.grid(row=2, column=0, sticky="ew")
        for col in range(4):
            cards.grid_columnconfigure(col, weight=1)
        self._card(cards, 0, "♟", len(clientes), "Clientes", "Cadastrados")
        self._card(cards, 1, "✉", len(relatorios), "E-mails", "Enviados")
        self._card(cards, 2, "▥", len(relatorios), "Relatórios", "Registados")
        self._card(cards, 3, "♟♟", len(cc), "Destinatários CC", "Ativos")

        recente = ctk.CTkFrame(frame, fg_color=self.PANEL, corner_radius=12)
        recente.grid(row=3, column=0, sticky="ew", pady=(12, 8))
        recente.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(recente, text="Atividade recente", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, sticky="w", padx=18, pady=(14, 6))
        holder, tree = self._criar_tree(recente, ("Data/Hora", "Cliente", "E-mail", "Status", "Mensagem"), [145, 150, 210, 90, 300], 3)
        holder.grid(row=1, column=0, sticky="ew", padx=12, pady=(0, 10))
        for rel in relatorios[:5]:
            tree.insert("", "end", values=(rel.data_envio.strftime("%d/%m/%Y %H:%M"), rel.nome, rel.email, rel.status, rel.mensagem))

        acoes = ctk.CTkFrame(frame, fg_color=self.PANEL, corner_radius=12)
        acoes.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        ctk.CTkLabel(acoes, text="Ações rápidas", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=18, pady=(14, 8))
        linha = ctk.CTkFrame(acoes, fg_color="transparent")
        linha.pack(fill="x", padx=12, pady=(0, 16))
        for texto, pagina in [
            ("♟ Novo Cliente", self.pagina_clientes),
            ("✉ Enviar E-mails", self.pagina_envio),
            ("≡ Relatório", self.pagina_relatorios),
            ("⚙ Corpo do E-mail", self.pagina_corpo),
            ("♟ Gerir CC", self.pagina_cc),
        ]:
            ctk.CTkButton(linha, text=texto, height=72, fg_color=self.PANEL_2, hover_color="#3B3F44", command=pagina).pack(side="left", expand=True, fill="x", padx=6)

    def pagina_clientes(self):
        frame = self.pagina("Cadastro de Clientes")

        form = ctk.CTkFrame(
            frame,
            fg_color=self.PANEL,
            corner_radius=12
        )

        form.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        form.grid_columnconfigure(1, weight=1)

        entradas = {}

        labels = [
            "CIL",
            "Nome",
            "E-mail",
            "Arquivo anexo"
        ]

        for row, label in enumerate(labels):
            ctk.CTkLabel(
                form,
                text=label,
                font=ctk.CTkFont(weight="bold")
            ).grid(
                row=row,
                column=0,
                padx=18,
                pady=10,
                sticky="w"
            )

            entrada = ctk.CTkEntry(
                form,
                height=38
            )

            if label == "Arquivo anexo":
                entrada.configure(state="disabled")

            entrada.grid(
                row=row,
                column=1,
                padx=18,
                pady=10,
                sticky="ew"
            )

            entradas[label] = entrada

        # O campo do anexo será preenchido automaticamente.
        entradas["Arquivo anexo"].configure(
            state="disabled",
            placeholder_text="Gerado automaticamente pelo CIL"
        )

        def atualizar_nome_anexo(event=None):
            cil = entradas["CIL"].get().strip()

            entradas["Arquivo anexo"].configure(state="normal")
            entradas["Arquivo anexo"].delete(0, "end")

            if cil:
                entradas["Arquivo anexo"].insert(
                    0,
                    f"{cil}.pdf"
                )

            entradas["Arquivo anexo"].configure(state="disabled")

        # Atualiza automaticamente quando o utilizador escreve no CIL.
        entradas["CIL"].bind(
            "<KeyRelease>",
            atualizar_nome_anexo
        )

        def limpar_campos():
            for nome, entrada in entradas.items():
                entrada.configure(state="normal")
                entrada.delete(0, "end")

                if nome == "Arquivo anexo":
                    entrada.configure(state="disabled")

            entradas["CIL"].focus_set()

        def guardar():
            cil = entradas["CIL"].get().strip()
            nome = entradas["Nome"].get().strip()
            email = entradas["E-mail"].get().strip()

            # Nome do PDF criado automaticamente.
            arquivo_anexo = "{}.pdf".format(cil)

            if not cil:
                messagebox.showwarning(
                    "Validação",
                    "Informe o CIL do cliente."
                )
                entradas["CIL"].focus_set()
                return

            if not cil.isdigit():
                messagebox.showwarning(
                    "Validação",
                    "O CIL deve conter apenas números."
                )
                entradas["CIL"].focus_set()
                return

            if not nome:
                messagebox.showwarning(
                    "Validação",
                    "Informe o nome do cliente."
                )
                entradas["Nome"].focus_set()
                return

            if not EMAIL_RE.match(email):
                messagebox.showwarning(
                    "Validação",
                    "Informe um endereço de e-mail válido."
                )
                entradas["E-mail"].focus_set()
                return

            cliente = Cliente(
                cil=cil,
                nome=nome,
                email=email,
                arquivo_anexo=arquivo_anexo
            )

            try:
                self.clientes.inserir(cliente)

                self.aud.log(
                    self.usuario.id,
                    "INSERIR",
                    "Clientes",
                    "Cliente {} inserido com o anexo {}.".format(
                        cliente.cil,
                        cliente.arquivo_anexo
                    )
                )

                messagebox.showinfo(
                    "Sucesso",
                    "Cliente guardado com sucesso.\n\n"
                    "Arquivo associado: {}".format(
                        arquivo_anexo
                    )
                )

                limpar_campos()

            except Exception as exc:
                messagebox.showerror(
                    "Erro",
                    str(exc)
                )

        botoes = ctk.CTkFrame(
            frame,
            fg_color="transparent"
        )

        botoes.grid(
            row=2,
            column=0,
            sticky="e",
            pady=14
        )

        ctk.CTkButton(
            botoes,
            text="Guardar",
            width=130,
            command=guardar
        ).pack(
            side="left",
            padx=5
        )

        ctk.CTkButton(
            botoes,
            text="Limpar",
            width=130,
            fg_color="#3B3F44",
            hover_color="#4A4E53",
            command=limpar_campos
        ).pack(
            side="left",
            padx=5
        )

    def pagina_consulta(self):
        frame = self.pagina("Consulta e CRUD de Clientes", scroll=False)
        busca = ctk.CTkEntry(frame, placeholder_text="Filtrar por CIL, nome ou e-mail", height=38)
        busca.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        holder, tree = self._criar_tree(frame, ("CIL", "Nome", "E-mail", "Anexo"), [125, 190, 240, 180], 12)
        holder.grid(row=2, column=0, sticky="nsew")
        frame.grid_rowconfigure(2, weight=1)
        barra = ctk.CTkFrame(frame, fg_color="transparent")
        barra.grid(row=3, column=0, sticky="ew", pady=12)
        novo_email = ctk.CTkEntry(barra, placeholder_text="Novo e-mail", width=280)
        novo_email.pack(side="left")
        cache = {}

        def carregar(*_):
            cache.clear()
            tree.delete(*tree.get_children())
            termo = busca.get().strip().lower()
            try:
                for cliente in self.clientes.listar():
                    if termo in "{} {} {}".format(cliente.cil, cliente.nome, cliente.email).lower():
                        item = tree.insert("", "end", values=(cliente.cil, cliente.nome, cliente.email, cliente.arquivo_anexo))
                        cache[item] = cliente
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        def atualizar_email():
            selecao = tree.selection()
            email = novo_email.get().strip()
            if not selecao or not EMAIL_RE.match(email):
                messagebox.showwarning("Validação", "Selecione um cliente e introduza um e-mail válido.")
                return
            cliente = cache[selecao[0]]
            try:
                self.clientes.atualizar(cliente.cil, Cliente(cliente.cil, cliente.nome, email, cliente.arquivo_anexo))
                carregar()
                novo_email.delete(0, "end")
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        def eliminar():
            selecao = tree.selection()
            if not selecao:
                return
            cliente = cache[selecao[0]]
            if messagebox.askyesno("Confirmar", "Eliminar o cliente {}?".format(cliente.cil)):
                try:
                    self.clientes.eliminar(cliente.cil)
                    carregar()
                except Exception as exc:
                    messagebox.showerror("Erro", str(exc))

        ctk.CTkButton(barra, text="Atualizar e-mail", command=atualizar_email).pack(side="left", padx=10)
        ctk.CTkButton(barra, text="Eliminar", command=eliminar).pack(side="left")
        busca.bind("<KeyRelease>", carregar)
        carregar()

    def pagina_envio(self):
        frame = self.pagina("Enviar E-mails", scroll=False)
        topo = ctk.CTkFrame(frame, fg_color=self.PANEL, corner_radius=12)
        topo.grid(row=1, column=0, sticky="ew")
        topo.grid_columnconfigure(0, weight=1)
        pasta = ctk.StringVar()
        ctk.CTkEntry(topo, textvariable=pasta, placeholder_text="Pasta dos ficheiros PDF", height=38).grid(row=0, column=0, padx=14, pady=14, sticky="ew")
        ctk.CTkButton(topo, text="Escolher pasta", command=lambda: pasta.set(filedialog.askdirectory())).grid(row=0, column=1, padx=(0, 14))
        progresso = ctk.CTkProgressBar(frame)
        progresso.grid(row=2, column=0, sticky="ew", pady=14)
        progresso.set(0)
        log = ctk.CTkTextbox(frame, height=390)
        log.grid(row=3, column=0, sticky="nsew")
        frame.grid_rowconfigure(3, weight=1)

        def executar():
            try:
                clientes = list(self.clientes.listar())
                corpo = self.conf.obter_corpo_email()
                destinatarios_cc = list(self.cc.listar())
            except Exception as exc:
                self.after(0, lambda: messagebox.showerror("Erro", str(exc)))
                return
            total = max(len(clientes), 1)
            for index, cliente in enumerate(clientes, 1):
                nome_anexo = (
                    cliente.arquivo_anexo
                    or "{}.pdf".format(cliente.cil)
                )

                anexo = Path(pasta.get()) / nome_anexo

                try:
                    if not anexo.is_file():
                        raise FileNotFoundError(
                            "O ficheiro PDF não foi encontrado: {}".format(
                                anexo
                            )
                        )

                    assunto = self.config_service.get(
                        "EMAIL",
                        "ASSUNTO",
                        "Factura de Energia"
                    )

                    self.email.enviar(
                        destinatario=cliente.email,
                        assunto=assunto,
                        corpo=corpo,
                        anexo=str(anexo),
                        cc=destinatarios_cc
                    )

                    status = "SUCESSO"
                    mensagem = "E-mail enviado com sucesso."

                except Exception as exc:
                    status = "ERRO"
                    mensagem = str(exc)
                try:
                    self.relatorios.inserir(RelatorioEnvio(cliente.nome, cliente.email, cliente.cil, status, mensagem, datetime.now()))
                except Exception:
                    pass
                self.after(0, lambda c=cliente, s=status, m=mensagem, i=index: (
                    log.insert("end", "{} - {}: {}\n".format(c.cil, s, m)),
                    log.see("end"),
                    progresso.set(i / total),
                ))
            self.after(0, lambda: messagebox.showinfo("Concluído", "Processamento terminado."))

        def iniciar():
            if not Path(pasta.get()).is_dir():
                messagebox.showwarning("Pasta", "Selecione uma pasta válida.")
                return
            threading.Thread(target=executar, daemon=True).start()

        ctk.CTkButton(frame, text="Iniciar envio", width=140, command=iniciar).grid(row=4, column=0, sticky="e", pady=12)

    def pagina_corpo(self):
        frame = self.pagina("Corpo do E-mail", scroll=False)
        texto = ctk.CTkTextbox(frame, height=520)
        texto.grid(row=1, column=0, sticky="nsew")
        frame.grid_rowconfigure(1, weight=1)
        try:
            texto.insert("1.0", self.conf.obter_corpo_email())
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

        def guardar():
            try:
                self.conf.guardar_corpo_email(texto.get("1.0", "end-1c"))
                messagebox.showinfo("Sucesso", "Corpo do e-mail guardado.")
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        ctk.CTkButton(frame, text="Guardar", width=130, command=guardar).grid(row=2, column=0, sticky="e", pady=12)

    def pagina_cc(self):
        frame = self.pagina("Gestão de Destinatários CC", scroll=False)
        entrada = ctk.CTkEntry(frame, placeholder_text="email@dominio.cv", height=38)
        entrada.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        holder, tree = self._criar_tree(frame, ("E-mail CC",), [600], 13)
        holder.grid(row=2, column=0, sticky="nsew")
        frame.grid_rowconfigure(2, weight=1)
        barra = ctk.CTkFrame(frame, fg_color="transparent")
        barra.grid(row=3, column=0, sticky="e", pady=12)

        def carregar():
            tree.delete(*tree.get_children())
            try:
                for email in self.cc.listar():
                    tree.insert("", "end", values=(email,))
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        def adicionar():
            email = entrada.get().strip()
            if not EMAIL_RE.match(email):
                messagebox.showwarning("Validação", "E-mail inválido.")
                return
            try:
                self.cc.inserir(email)
                entrada.delete(0, "end")
                carregar()
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        def eliminar():
            selecao = tree.selection()
            email = tree.item(selecao[0], "values")[0] if selecao else entrada.get().strip()
            if not email:
                return
            try:
                self.cc.eliminar(email)
                carregar()
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))

        ctk.CTkButton(barra, text="Adicionar", command=adicionar).pack(side="left", padx=5)
        ctk.CTkButton(barra, text="Eliminar", command=eliminar).pack(side="left", padx=5)
        carregar()

    @staticmethod
    def _parse_data(valor):
        valor = valor.strip()
        if not valor:
            return None
        return datetime.strptime(valor, "%d/%m/%Y").date()

    def pagina_relatorios(self):
        frame = self.pagina("Relatório de Envios", scroll=False)
        filtros = ctk.CTkFrame(frame, fg_color="transparent")
        filtros.grid(row=1, column=0, sticky="w", pady=(0, 10))
        ctk.CTkLabel(filtros, text="Início:").pack(side="left", padx=(0, 6))
        inicio = ctk.CTkEntry(filtros, placeholder_text="dd/mm/aaaa", width=150)
        inicio.pack(side="left", padx=(0, 16))
        ctk.CTkLabel(filtros, text="Fim:").pack(side="left", padx=(0, 6))
        fim = ctk.CTkEntry(filtros, placeholder_text="dd/mm/aaaa", width=150)
        fim.pack(side="left", padx=(0, 16))
        holder, tree = self._criar_tree(frame, ("Data", "Nome", "Email", "CIL", "Status", "Mensagem"), [140, 150, 200, 100, 90, 310], 12)
        holder.grid(row=2, column=0, sticky="nsew")
        frame.grid_rowconfigure(2, weight=1)
        cache: List[RelatorioEnvio] = []

        def carregar():
            try:
                data_inicio = self._parse_data(inicio.get())
                data_fim = self._parse_data(fim.get())
            except ValueError:
                messagebox.showwarning("Data", "Utilize o formato dd/mm/aaaa.")
                return
            tree.delete(*tree.get_children())
            cache[:] = list(self.relatorios.listar(data_inicio, data_fim))
            for rel in cache:
                tree.insert("", "end", values=(rel.data_envio.strftime("%d/%m/%Y %H:%M"), rel.nome, rel.email, rel.cil, rel.status, rel.mensagem))

        ctk.CTkButton(filtros, text="Filtrar", command=carregar).pack(side="left")
        barra = ctk.CTkFrame(frame, fg_color="transparent")
        barra.grid(row=3, column=0, sticky="w", pady=12)

        def excel():
            caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
            if caminho:
                exportar_excel(cache, caminho)
                messagebox.showinfo("Sucesso", "Relatório Excel exportado.")

        def pdf():
            caminho = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
            if caminho:
                exportar_pdf(cache, caminho)
                messagebox.showinfo("Sucesso", "Relatório PDF exportado.")

        ctk.CTkButton(barra, text="Exportar Excel", command=excel).pack(side="left", padx=(0, 10))
        ctk.CTkButton(barra, text="Exportar PDF", command=pdf).pack(side="left")
        carregar()

    def pagina_usuarios(self):
        frame = self.pagina("Gestão de Utilizadores", scroll=False)

        ctk.CTkLabel(
            frame,
            text="Crie e altere utilizadores, e-mails, palavras-passe e níveis de acesso."
        ).grid(row=1, column=0, sticky="w", pady=(0, 12))

        form = ctk.CTkFrame(frame, fg_color="transparent")
        form.grid(row=2, column=0, sticky="ew")
        form.grid_columnconfigure(1, weight=1)

        campos = {}

        for row, nome in enumerate(
            ("Utilizador", "Email", "Palavra-passe", "Confirmar")
        ):
            ctk.CTkLabel(
                form,
                text="{}:".format(nome)
            ).grid(
                row=row,
                column=0,
                sticky="w",
                pady=6,
                padx=(0, 14)
            )

            ocultar = nome in ("Palavra-passe", "Confirmar")

            entrada = ctk.CTkEntry(
                form,
                show="•" if ocultar else "",
                height=36
            )
            entrada.grid(
                row=row,
                column=1,
                sticky="ew",
                pady=6
            )
            campos[nome] = entrada

        ctk.CTkLabel(
            form,
            text="Nível:"
        ).grid(
            row=4,
            column=0,
            sticky="w",
            pady=6,
            padx=(0, 14)
        )

        nivel = ctk.CTkComboBox(
            form,
            values=["admin", "gerente"],
            state="readonly",
            height=36
        )
        nivel.grid(
            row=4,
            column=1,
            sticky="ew",
            pady=6
        )
        nivel.set("gerente")

        mostrar = ctk.BooleanVar(value=False)

        def alternar():
            mascara = "" if mostrar.get() else "•"
            campos["Palavra-passe"].configure(show=mascara)
            campos["Confirmar"].configure(show=mascara)

        ctk.CTkCheckBox(
            form,
            text="Mostrar palavras-passe",
            variable=mostrar,
            command=alternar
        ).grid(
            row=5,
            column=1,
            sticky="w",
            pady=6
        )

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(
            row=3,
            column=0,
            sticky="w",
            pady=10
        )

        holder, tree = self._criar_tree(
            frame,
            ("ID", "Utilizador", "Email", "Palavra-passe", "Nível"),
            [70, 190, 260, 170, 130],
            9
        )
        holder.grid(
            row=4,
            column=0,
            sticky="nsew"
        )
        frame.grid_rowconfigure(4, weight=1)

        cache = {}
        selecionado = {"id": None}

        def limpar():
            selecionado["id"] = None

            for entrada in campos.values():
                entrada.delete(0, "end")

            nivel.set("gerente")
            mostrar.set(False)
            alternar()

            selecoes = tree.selection()
            if selecoes:
                tree.selection_remove(*selecoes)

            campos["Utilizador"].focus_set()

        def carregar():
            cache.clear()
            tree.delete(*tree.get_children())

            try:
                for usuario in self.usuarios.listar():
                    email_usuario = getattr(usuario, "email", "") or ""
                    password_usuario = getattr(usuario, "password", "") or ""

                    item = tree.insert(
                        "",
                        "end",
                        values=(
                            usuario.id,
                            usuario.username,
                            email_usuario,
                            "•" * 8 if password_usuario else "",
                            usuario.nivel
                        )
                    )
                    cache[item] = usuario

            except Exception as exc:
                messagebox.showerror(
                    "Erro",
                    "Não foi possível carregar os utilizadores.\n\n{}".format(exc)
                )

        def selecionar(_event=None):
            selecao = tree.selection()

            if not selecao:
                return

            usuario = cache.get(selecao[0])

            if usuario is None:
                return

            selecionado["id"] = usuario.id

            campos["Utilizador"].delete(0, "end")
            campos["Utilizador"].insert(
                0,
                getattr(usuario, "username", "") or ""
            )

            campos["Email"].delete(0, "end")
            campos["Email"].insert(
                0,
                getattr(usuario, "email", "") or ""
            )

            password_usuario = getattr(usuario, "password", "") or ""

            campos["Palavra-passe"].delete(0, "end")
            campos["Palavra-passe"].insert(0, password_usuario)

            campos["Confirmar"].delete(0, "end")
            campos["Confirmar"].insert(0, password_usuario)

            nivel.set(
                getattr(usuario, "nivel", "gerente") or "gerente"
            )

        def validar():
            username = campos["Utilizador"].get().strip()
            email_usuario = campos["Email"].get().strip()
            password = campos["Palavra-passe"].get()
            confirmacao = campos["Confirmar"].get()

            if not username:
                messagebox.showwarning(
                    "Validação",
                    "Informe o utilizador."
                )
                campos["Utilizador"].focus_set()
                return False

            if not email_usuario:
                messagebox.showwarning(
                    "Validação",
                    "Informe o e-mail do utilizador."
                )
                campos["Email"].focus_set()
                return False

            if not EMAIL_RE.match(email_usuario):
                messagebox.showwarning(
                    "Validação",
                    "Informe um endereço de e-mail válido."
                )
                campos["Email"].focus_set()
                return False

            if not password:
                messagebox.showwarning(
                    "Validação",
                    "Informe a palavra-passe."
                )
                campos["Palavra-passe"].focus_set()
                return False

            if password != confirmacao:
                messagebox.showwarning(
                    "Validação",
                    "As palavras-passe não coincidem."
                )
                campos["Confirmar"].focus_set()
                return False

            return True

        def guardar():
            if not validar():
                return

            username = campos["Utilizador"].get().strip()
            email_usuario = campos["Email"].get().strip()
            password = campos["Palavra-passe"].get()
            nivel_usuario = nivel.get()

            try:
                self.usuarios.inserir(
                    username,
                    email_usuario,
                    password,
                    nivel_usuario
                )

                carregar()
                limpar()

                messagebox.showinfo(
                    "Sucesso",
                    "Utilizador guardado com sucesso."
                )

            except Exception as exc:
                messagebox.showerror(
                    "Erro",
                    str(exc)
                )

        def alterar():
            if selecionado["id"] is None:
                messagebox.showwarning(
                    "Seleção",
                    "Selecione um utilizador."
                )
                return

            if not validar():
                return

            username = campos["Utilizador"].get().strip()
            email_usuario = campos["Email"].get().strip()
            password = campos["Palavra-passe"].get()
            nivel_usuario = nivel.get()

            try:
                self.usuarios.atualizar(
                    selecionado["id"],
                    username,
                    email_usuario,
                    password,
                    nivel_usuario
                )

                carregar()
                limpar()

                messagebox.showinfo(
                    "Sucesso",
                    "Utilizador alterado com sucesso."
                )

            except Exception as exc:
                messagebox.showerror(
                    "Erro",
                    str(exc)
                )

        def eliminar():
            if selecionado["id"] is None:
                messagebox.showwarning(
                    "Seleção",
                    "Selecione um utilizador."
                )
                return

            if selecionado["id"] == self.usuario.id:
                messagebox.showwarning(
                    "Operação",
                    "Não pode eliminar o utilizador atualmente autenticado."
                )
                return

            if not messagebox.askyesno(
                "Confirmar",
                "Eliminar o utilizador selecionado?"
            ):
                return

            try:
                self.usuarios.eliminar(
                    selecionado["id"]
                )

                carregar()
                limpar()

                messagebox.showinfo(
                    "Sucesso",
                    "Utilizador eliminado com sucesso."
                )

            except Exception as exc:
                messagebox.showerror(
                    "Erro",
                    str(exc)
                )

        ctk.CTkButton(
            botoes,
            text="Novo",
            command=limpar
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            botoes,
            text="Guardar",
            command=guardar
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            botoes,
            text="Alterar",
            command=alterar
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            botoes,
            text="Eliminar",
            command=eliminar
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            botoes,
            text="Limpar",
            command=limpar
        ).pack(side="left", padx=8)

        tree.bind(
            "<<TreeviewSelect>>",
            selecionar
        )

        carregar()

    def pagina_definicoes(self):
        frame = self.pagina("Definições")

        ctk.CTkLabel(
            frame,
            text="Nesta área pode alterar a palavra-passe da sua conta.",
            font=ctk.CTkFont(size=15),
        ).grid(row=1, column=0, sticky="w", pady=(0, 16))

        card = ctk.CTkFrame(
            frame,
            fg_color=self.PANEL,
            corner_radius=12,
            border_width=1,
            border_color=self.BORDER,
        )
        card.grid(row=2, column=0, sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        campos = {}
        dados = [
            ("Palavra-passe atual", "Introduza a palavra-passe atual"),
            ("Nova palavra-passe", "Introduza a nova palavra-passe"),
            ("Confirmar nova palavra-passe", "Repita a nova palavra-passe"),
        ]

        for linha, (rotulo, placeholder) in enumerate(dados):
            ctk.CTkLabel(
                card,
                text=rotulo,
                font=ctk.CTkFont(weight="bold"),
            ).grid(row=linha, column=0, padx=18, pady=12, sticky="w")

            entrada = ctk.CTkEntry(
                card,
                placeholder_text=placeholder,
                show="•",
                height=40,
            )
            entrada.grid(row=linha, column=1, padx=18, pady=12, sticky="ew")
            campos[rotulo] = entrada

        mostrar = ctk.BooleanVar(value=False)

        def alternar_visibilidade():
            caractere = "" if mostrar.get() else "•"
            for entrada in campos.values():
                entrada.configure(show=caractere)

        ctk.CTkCheckBox(
            card,
            text="Mostrar palavras-passe",
            variable=mostrar,
            command=alternar_visibilidade,
        ).grid(row=3, column=1, padx=18, pady=(2, 12), sticky="w")

        def limpar_campos():
            for entrada in campos.values():
                entrada.delete(0, "end")
            campos["Palavra-passe atual"].focus_set()

        def alterar_palavra_passe():
            atual = campos["Palavra-passe atual"].get()
            nova = campos["Nova palavra-passe"].get()
            confirmar = campos["Confirmar nova palavra-passe"].get()

            if not atual or not nova or not confirmar:
                messagebox.showwarning(
                    "Validação",
                    "Preencha todos os campos.",
                    parent=self,
                )
                return

            if len(nova) < 8:
                messagebox.showwarning(
                    "Validação",
                    "A nova palavra-passe deve ter pelo menos 8 caracteres.",
                    parent=self,
                )
                campos["Nova palavra-passe"].focus_set()
                return

            if nova != confirmar:
                messagebox.showwarning(
                    "Validação",
                    "A nova palavra-passe e a confirmação não coincidem.",
                    parent=self,
                )
                campos["Confirmar nova palavra-passe"].focus_set()
                return

            if atual == nova:
                messagebox.showwarning(
                    "Validação",
                    "A nova palavra-passe deve ser diferente da atual.",
                    parent=self,
                )
                return

            try:
                sucesso, mensagem = self.usuarios.alterar_palavra_passe(
                    self.usuario.id,
                    atual,
                    nova,
                )

                if not sucesso:
                    messagebox.showerror("Erro", mensagem, parent=self)
                    campos["Palavra-passe atual"].focus_set()
                    return

                self.usuario.password = nova
                self.aud.log(
                    self.usuario.id,
                    "ALTERAR_PALAVRA_PASSE",
                    "Definições",
                    "O utilizador alterou a palavra-passe da própria conta.",
                )
                limpar_campos()
                messagebox.showinfo("Sucesso", mensagem, parent=self)

            except Exception as exc:
                messagebox.showerror("Erro", str(exc), parent=self)

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.grid(row=3, column=0, sticky="e", pady=14)

        ctk.CTkButton(
            botoes,
            text="Alterar palavra-passe",
            width=190,
            command=alterar_palavra_passe,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            botoes,
            text="Limpar",
            width=120,
            fg_color="#3B3F44",
            hover_color="#4A4E53",
            command=limpar_campos,
        ).pack(side="left", padx=6)

        ctk.CTkLabel(
            frame,
            text="As ligações à base de dados e ao SMTP continuam configuradas no ficheiro config.ini.",
            text_color="#AEB6C2",
            font=ctk.CTkFont(size=13),
        ).grid(row=4, column=0, sticky="w", pady=(8, 0))
