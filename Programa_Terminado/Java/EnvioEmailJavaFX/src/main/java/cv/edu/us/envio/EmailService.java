package cv.edu.us.envio;

import jakarta.mail.*;
import jakarta.mail.internet.*;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Properties;

public final class EmailService {
    private final AppConfig config;

    public EmailService(AppConfig config) {
        this.config = config;
    }

    public void enviar(
            Cliente cliente,
            File anexo,
            List<String> cc,
            String corpoPadrao
    ) throws Exception {
        Properties props = new Properties();
        props.put("mail.smtp.host", config.get("email.smtpHost"));
        props.put("mail.smtp.port", config.get("email.smtpPort"));
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.ssl.enable", "true");

        Session session = Session.getInstance(props, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(
                    config.get("email.remetente"),
                    config.get("email.senhaApp")
                );
            }
        });

        MimeMessage msg = new MimeMessage(session);
        msg.setFrom(new InternetAddress(config.get("email.remetente")));
        msg.setRecipients(Message.RecipientType.TO, InternetAddress.parse(cliente.email()));

        for (String emailCc : cc) {
            if (!emailCc.isBlank()) {
                msg.addRecipient(Message.RecipientType.CC, new InternetAddress(emailCc));
            }
        }

        msg.setSubject(config.get("email.assunto"), StandardCharsets.UTF_8.name());

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
        msg.setContent(multipart);

        Transport.send(msg);
    }
}
