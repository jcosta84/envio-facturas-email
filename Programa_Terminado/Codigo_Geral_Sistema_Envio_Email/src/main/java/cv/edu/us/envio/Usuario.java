package cv.edu.us.envio;

/**
 * Representa um utilizador do sistema.
 *
 * A palavra-passe é incluída porque a versão atual da base
 * de dados guarda este valor em texto simples.
 */
public record Usuario(
        int id,
        String username,
        String password,
        String nivel
) {
    public Usuario {
        username =
            username == null
                ? ""
                : username.trim();

        password =
            password == null
                ? ""
                : password;

        nivel =
            nivel == null
                ? ""
                : nivel.trim()
                       .toLowerCase();
    }

    public boolean isAdmin() {
        return "admin".equalsIgnoreCase(nivel)
            || "administrador"
                .equalsIgnoreCase(nivel);
    }

    public boolean isGerente() {
        return "gerente"
            .equalsIgnoreCase(nivel);
    }
}
