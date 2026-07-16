package cv.edu.us.envio;

/**
 * Representa o utilizador autenticado no sistema.
 */
public record Usuario(
        int id,
        String username,
        String nivel
) {
    public Usuario {
        username = username == null ? "" : username.trim();
        nivel = nivel == null ? "" : nivel.trim().toLowerCase();
    }

    public boolean isAdmin() {
        return "admin".equalsIgnoreCase(nivel)
            || "administrador".equalsIgnoreCase(nivel);
    }

    public boolean isGerente() {
        return "gerente".equalsIgnoreCase(nivel);
    }
}
