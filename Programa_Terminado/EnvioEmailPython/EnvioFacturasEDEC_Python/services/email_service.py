import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable, List, Optional


class EmailService:
    def __init__(self, config_service):
        self.config = config_service

    def enviar(
        self,
        destinatario: str,
        corpo: str,
        anexo: str,
        cc: Optional[Iterable[str]] = None,
        assunto: Optional[str] = None
    ) -> None:
        """
        Envia uma mensagem com anexo PDF.

        Os parâmetros devem ser preferencialmente passados pelo nome para
        evitar erros na ordem dos argumentos.
        """

        if assunto is None:
            assunto = self.config.get(
                "EMAIL",
                "ASSUNTO",
                "Factura de Energia"
            )

        self.enviar_email(
            destinatario=destinatario,
            assunto=assunto,
            corpo=corpo,
            anexo=anexo,
            cc=cc
        )

    def enviar_email(
        self,
        destinatario: str,
        assunto: str,
        corpo: str,
        anexo: str,
        cc: Optional[Iterable[str]] = None
    ) -> None:

        # Garantir que os campos textuais são realmente strings
        destinatario = str(destinatario).strip()
        assunto = str(assunto).strip()
        corpo = str(corpo)
        anexo = str(anexo)

        remetente = self.config.get(
            "EMAIL",
            "REMETENTE"
        ).strip()

        username = self.config.get(
            "EMAIL",
            "USERNAME"
        ).strip()

        password = self.config.get(
            "EMAIL",
            "PASSWORD"
        ).strip()

        smtp_host = self.config.get(
            "EMAIL",
            "SMTP_HOST"
        ).strip()

        smtp_port = self.config.getint(
            "EMAIL",
            "SMTP_PORT",
            25
        )

        seguranca = self.config.get(
            "EMAIL",
            "SMTP_SEGURANCA",
            "STARTTLS"
        ).strip().upper()

        nome_remetente = self.config.get(
            "EMAIL",
            "NOME_REMETENTE",
            "EDEC SUL"
        ).strip()

        if not destinatario:
            raise ValueError(
                "O endereço do destinatário está vazio."
            )

        if not remetente:
            raise ValueError(
                "O endereço do remetente está vazio."
            )

        caminho_anexo = Path(anexo)

        if not caminho_anexo.is_file():
            raise FileNotFoundError(
                "O ficheiro PDF não foi encontrado:\n{}".format(
                    caminho_anexo
                )
            )

        lista_cc = self._normalizar_cc(cc)

        mensagem = EmailMessage()

        mensagem["From"] = "{} <{}>".format(
            nome_remetente,
            remetente
        )

        mensagem["To"] = destinatario
        mensagem["Subject"] = assunto

        if lista_cc:
            mensagem["Cc"] = ", ".join(lista_cc)

        mensagem.set_content(
            corpo,
            subtype="plain",
            charset="utf-8"
        )

        with caminho_anexo.open("rb") as ficheiro:
            conteudo_pdf = ficheiro.read()

        mensagem.add_attachment(
            conteudo_pdf,
            maintype="application",
            subtype="pdf",
            filename=caminho_anexo.name
        )

        contexto_ssl = ssl.create_default_context()

        if seguranca in ("SSL", "TLS_SSL"):
            self._enviar_ssl(
                mensagem=mensagem,
                host=smtp_host,
                porta=smtp_port,
                username=username,
                password=password,
                contexto=contexto_ssl
            )
        else:
            self._enviar_starttls(
                mensagem=mensagem,
                host=smtp_host,
                porta=smtp_port,
                username=username,
                password=password,
                contexto=contexto_ssl,
                usar_starttls=seguranca == "STARTTLS"
            )

    @staticmethod
    def _normalizar_cc(
        cc: Optional[Iterable[str]]
    ) -> List[str]:
        if cc is None:
            return []

        # Caso seja recebido apenas um endereço como texto
        if isinstance(cc, str):
            valores = cc.replace(";", ",").split(",")
        else:
            valores = list(cc)

        resultado = []

        for valor in valores:
            if valor is None:
                continue

            email = str(valor).strip()

            if email and email not in resultado:
                resultado.append(email)

        return resultado

    def _enviar_starttls(
        self,
        mensagem: EmailMessage,
        host: str,
        porta: int,
        username: str,
        password: str,
        contexto: ssl.SSLContext,
        usar_starttls: bool
    ) -> None:

        with smtplib.SMTP(
            host=host,
            port=porta,
            timeout=30
        ) as smtp:

            smtp.ehlo()

            if usar_starttls:
                if smtp.has_extn("STARTTLS"):
                    smtp.starttls(context=contexto)
                    smtp.ehlo()
                else:
                    print(
                        "Aviso: o servidor não disponibiliza STARTTLS."
                    )

            self._autenticar_se_disponivel(
                smtp,
                username,
                password
            )

            smtp.send_message(mensagem)

    def _enviar_ssl(
        self,
        mensagem: EmailMessage,
        host: str,
        porta: int,
        username: str,
        password: str,
        contexto: ssl.SSLContext
    ) -> None:

        with smtplib.SMTP_SSL(
            host=host,
            port=porta,
            context=contexto,
            timeout=30
        ) as smtp:

            smtp.ehlo()

            self._autenticar_se_disponivel(
                smtp,
                username,
                password
            )

            smtp.send_message(mensagem)

    @staticmethod
    def _autenticar_se_disponivel(
        smtp,
        username: str,
        password: str
    ) -> None:

        if not username or not password:
            return

        if smtp.has_extn("AUTH"):
            smtp.login(
                username,
                password
            )
        else:
            print(
                "Aviso: o servidor SMTP não oferece autenticação. "
                "O envio será tentado sem login."
            )