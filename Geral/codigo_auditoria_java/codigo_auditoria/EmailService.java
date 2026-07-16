package cv.edu.us.envio;

import jakarta.mail.Authenticator;
import jakarta.mail.Message;
import jakarta.mail.Multipart;
import jakarta.mail.PasswordAuthentication;
import jakarta.mail.Session;
import jakarta.mail.Transport;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeBodyPart;
import jakarta.mail.internet.MimeMessage;
import jakarta.mail.internet.MimeMultipart;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

public final class EmailService {

    private final ConfigIni config;

    public EmailService(ConfigIni config) {
        this.config = config;
    }

    public void enviar(
            Cliente cliente,
            File anexo,
            Iterable<String> emailsCc,
            String corpoPadrao
    ) throws Exception {

        String smtpHost = valorOuPadrao(
                config.get("EMAIL", "SMTP_HOST"),
                "mail.edec.cv"
        );

        String smtpPort = valorOuPadrao(
                config.get("EMAIL", "SMTP_PORT"),
                "25"
        );

        String seguranca = valorOuPadrao(
                config.get("EMAIL", "SMTP_SEGURANCA"),
                "STARTTLS"
        );

        String remetente = config.get(
                "EMAIL",
                "REMETENTE"
        );

        /*
         * O nome utilizado para autenticação pode ser diferente
         * do endereço de e-mail.
         *
         * Exemplo:
         * remetente: p.costa@edec.cv
         * utilizador: edec.pcosta
         */
        String username = valorOuPadrao(
                config.get("EMAIL", "USERNAME"),
                remetente
        );

        /*
         * Primeiro procura PASSWORD.
         * Caso não exista, utiliza SENHA_APP para manter
         * compatibilidade com a configuração anterior do Gmail.
         */
        String password = config.get(
                "EMAIL",
                "PASSWORD"
        );

        if (password == null || password.isBlank()) {
            password = config.get(
                    "EMAIL",
                    "SENHA_APP"
            );
        }

        String assunto = valorOuPadrao(
                config.get("EMAIL", "ASSUNTO"),
                "Factura de Energia da Empresa EDEC"
        );

        validarConfiguracao(
                remetente,
                username,
                password
        );

        validarAnexo(anexo);

        Properties props = criarPropriedadesSmtp(
                smtpHost,
                smtpPort,
                seguranca
        );

        String utilizadorFinal = username;
        String passwordFinal = password;

        Session session = Session.getInstance(
                props,
                new Authenticator() {
                    @Override
                    protected PasswordAuthentication
                    getPasswordAuthentication() {

                        return new PasswordAuthentication(
                                utilizadorFinal,
                                passwordFinal
                        );
                    }
                }
        );

        /*
         * Ative temporariamente para visualizar no terminal
         * todos os detalhes da comunicação SMTP.
         */
        // session.setDebug(true);

        MimeMessage mensagem = new MimeMessage(session);

        mensagem.setFrom(
                new InternetAddress(
                        remetente,
                        "EDEC SUL",
                        StandardCharsets.UTF_8.name()
                )
                );

        mensagem.setRecipients(
                Message.RecipientType.TO,
                InternetAddress.parse(
                        cliente.email(),
                        false
                )
        );

        adicionarEmailsCc(
                mensagem,
                emailsCc
        );

        mensagem.setSubject(
                assunto,
                StandardCharsets.UTF_8.name()
        );

        MimeBodyPart parteTexto = new MimeBodyPart();

        parteTexto.setText(
                criarCorpoMensagem(
                        cliente,
                        corpoPadrao
                ),
                StandardCharsets.UTF_8.name()
        );

        MimeBodyPart parteAnexo = new MimeBodyPart();
        parteAnexo.attachFile(anexo);

        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(parteTexto);
        multipart.addBodyPart(parteAnexo);

        mensagem.setContent(multipart);
        mensagem.saveChanges();

        Transport.send(mensagem);
    }

    private Properties criarPropriedadesSmtp(
            String smtpHost,
            String smtpPort,
            String seguranca
    ) {

        Properties props = new Properties();

        props.put(
                "mail.smtp.host",
                smtpHost
        );

        props.put(
                "mail.smtp.port",
                smtpPort
        );

        props.put(
                "mail.smtp.auth",
                "true"
        );

        /*
         * Tempos máximos em milissegundos.
         * Evitam que o programa fique bloqueado indefinidamente.
         */
        props.put(
                "mail.smtp.connectiontimeout",
                "15000"
        );

        props.put(
                "mail.smtp.timeout",
                "30000"
        );

        props.put(
                "mail.smtp.writetimeout",
                "30000"
        );

        if ("SSL".equalsIgnoreCase(seguranca)
                || "SSL/TLS".equalsIgnoreCase(seguranca)) {

            /*
             * Normalmente utilizado na porta 465.
             */
            props.put(
                    "mail.smtp.ssl.enable",
                    "true"
            );

            props.put(
                    "mail.smtp.starttls.enable",
                    "false"
            );

        } else if ("STARTTLS".equalsIgnoreCase(seguranca)
                || "TLS".equalsIgnoreCase(seguranca)) {

            /*
             * Configuração adequada ao servidor da EDEC:
             * mail.edec.cv, porta 25, STARTTLS.
             */
            props.put(
                    "mail.smtp.ssl.enable",
                    "false"
            );

            props.put(
                    "mail.smtp.starttls.enable",
                    "true"
            );

            /*
             * Em alguns servidores antigos, o STARTTLS pode ser
             * disponibilizado mas não ser obrigatório.
             *
             * Caso a ligação falhe, altere este valor para false.
             */
            props.put(
                    "mail.smtp.starttls.required",
                    "true"
            );

        } else {

            /*
             * Ligação sem encriptação.
             * Não é recomendada.
             */
            props.put(
                    "mail.smtp.ssl.enable",
                    "false"
            );

            props.put(
                    "mail.smtp.starttls.enable",
                    "false"
            );
        }

        return props;
    }

    private void adicionarEmailsCc(
            MimeMessage mensagem,
            Iterable<String> emailsCc
    ) throws Exception {

        if (emailsCc == null) {
            return;
        }

        for (String emailCc : emailsCc) {

            if (emailCc == null
                    || emailCc.isBlank()) {
                continue;
            }

            mensagem.addRecipient(
                    Message.RecipientType.CC,
                    new InternetAddress(
                            emailCc.trim(),
                            false
                    )
            );
        }
    }

    private String criarCorpoMensagem(
            Cliente cliente,
            String corpoPadrao
    ) {

        String corpo = corpoPadrao == null
                ? ""
                : corpoPadrao.trim();

        return "Olá " + cliente.nome() + ",\n\n"
                + "CIL: " + cliente.cil() + "\n\n"
                + corpo;
    }

    private void validarConfiguracao(
            String remetente,
            String username,
            String password
    ) {

        if (remetente == null
                || remetente.isBlank()) {

            throw new IllegalStateException(
                    "O REMETENTE não foi definido no config.ini."
            );
        }

        if (username == null
                || username.isBlank()) {

            throw new IllegalStateException(
                    "O USERNAME não foi definido no config.ini."
            );
        }

        if (password == null
                || password.isBlank()) {

            throw new IllegalStateException(
                    "A PASSWORD não foi definida no config.ini."
            );
        }
    }

    private void validarAnexo(File anexo) {

        if (anexo == null) {
            throw new IllegalArgumentException(
                    "O ficheiro anexo não foi informado."
            );
        }

        if (!anexo.exists()) {
            throw new IllegalArgumentException(
                    "O ficheiro anexo não existe: "
                            + anexo.getAbsolutePath()
            );
        }

        if (!anexo.isFile()) {
            throw new IllegalArgumentException(
                    "O caminho informado não é um ficheiro."
            );
        }
    }

    private String valorOuPadrao(
            String valor,
            String padrao
    ) {

        return valor == null
                || valor.isBlank()
                ? padrao
                : valor.trim();
    }
}