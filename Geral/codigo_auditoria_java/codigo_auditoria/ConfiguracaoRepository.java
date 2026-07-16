package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public final class ConfiguracaoRepository {

    private final Database db;

    public ConfiguracaoRepository(Database db) {
        this.db = db;
    }

    public String obterCorpoEmail() throws SQLException {
        String sql =
            "SELECT conteudo FROM corpo_email WHERE id=1";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql);
             ResultSet rs = stmt.executeQuery()) {

            return rs.next() ? rs.getString("conteudo") : "";
        }
    }

    public void guardarCorpoEmail(
            String conteudo
    ) throws SQLException {

        String sql = """
            INSERT INTO corpo_email(id, conteudo)
            VALUES(1, ?)
            ON DUPLICATE KEY UPDATE conteudo=VALUES(conteudo)
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, conteudo);
            stmt.executeUpdate();
        }
    }
}
