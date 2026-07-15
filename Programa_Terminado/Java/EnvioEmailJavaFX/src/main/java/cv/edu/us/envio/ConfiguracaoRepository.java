package cv.edu.us.envio;

import java.sql.*;

public final class ConfiguracaoRepository {
    private final Database db;

    public ConfiguracaoRepository(Database db) {
        this.db = db;
    }

    public String obterCorpoEmail() throws SQLException {
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement("SELECT conteudo FROM corpo_email WHERE id=1");
             ResultSet rs = ps.executeQuery()) {
            return rs.next() ? rs.getString(1) : "";
        }
    }

    public void guardarCorpoEmail(String conteudo) throws SQLException {
        String sql = """
            INSERT INTO corpo_email(id, conteudo) VALUES(1, ?)
            ON DUPLICATE KEY UPDATE conteudo=VALUES(conteudo)
            """;
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, conteudo);
            ps.executeUpdate();
        }
    }
}
