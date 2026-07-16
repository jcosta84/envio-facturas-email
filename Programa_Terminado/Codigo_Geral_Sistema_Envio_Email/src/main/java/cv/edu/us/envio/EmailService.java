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

        String smtpHost =
            valorOuPadrao(
                config.get("EMAIL", "SMTP_HOST"),
                "smtp.gmail.com"
            );

        String smtpPort =
            valorOuPadrao(
                config.get("EMAIL", "SMTP_PORT"),
                "465"
            );

        String remetente =
            config.get("EMAIL", "REMETENTE");

        String senhaApp =
            config.get("EMAIL", "SENHA_APP");

        String assunto =
            config.get("EMAIL", "ASSUNTO");

        Properties props = new Properties();
        props.put("mail.smtp.host", smtpHost);
        props.put("mail.smtp.port", smtpPort);
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.ssl.enable", "true");

        Session session = Session.getInstance(
            props,
            new Authenticator() {
                @Override
                protected PasswordAuthentication
                getPasswordAuthentication() {

                    return new PasswordAuthentication(
                        remetente,
                        senhaApp
                    );
                }
            }
        );

        MimeMessage mensagem = new MimeMessage(session);

        mensagem.setFrom(
            new InternetAddress(remetente)
        );

        mensagem.setRecipients(
            Message.RecipientType.TO,
            InternetAddress.parse(cliente.email())
        );

        for (String emailCc : emailsCc) {
            if (emailCc != null && !emailCc.isBlank()) {
                mensagem.addRecipient(
                    Message.RecipientType.CC,
                    new InternetAddress(emailCc.trim())
                );
            }
        }

        mensagem.setSubject(
            assunto,
            StandardCharsets.UTF_8.name()
        );

        MimeBodyPart texto = new MimeBodyPart();

        texto.setText(
            "Olá " + cliente.nome() + ",\n\n" +
            "CIL: " + cliente.cil() + "\n\n" +
            corpoPadrao,
            StandardCharsets.UTF_8.name()
        );

        MimeBodyPart ficheiro = new MimeBodyPart();
        ficheiro.attachFile(anexo);

        Multipart multipart = new MimeMultipart();
        multipart.addBodyPart(texto);
        multipart.addBodyPart(ficheiro);

        mensagem.setContent(multipart);

        Transport.send(mensagem);
    }

    private String valorOuPadrao(
            String valor,
            String padrao
    ) {
        return valor == null || valor.isBlank()
            ? padrao
            : valor.trim();
    }
}
