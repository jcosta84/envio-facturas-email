package cv.edu.us.envio;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public final class AppConfig {
    private final Properties props = new Properties();

    public AppConfig() {
        try (InputStream in = getClass().getResourceAsStream("/config.properties")) {
            if (in == null) {
                throw new IllegalStateException("Ficheiro config.properties não encontrado.");
            }
            props.load(in);
            validar();
        } catch (IOException e) {
            throw new IllegalStateException("Erro ao carregar configurações.", e);
        }
    }

    private void validar() {
        String[] obrigatorios = {
            "db.host", "db.port", "db.name", "db.user", "db.password",
            "email.remetente", "email.senhaApp", "email.assunto",
            "email.smtpHost", "email.smtpPort"
        };
        for (String chave : obrigatorios) {
            if (get(chave).isBlank()) {
                throw new IllegalStateException("Configuração obrigatória vazia: " + chave);
            }
        }
    }

    public String get(String chave) {
        return props.getProperty(chave, "").trim();
    }

    public int getInt(String chave) {
        return Integer.parseInt(get(chave));
    }
}
