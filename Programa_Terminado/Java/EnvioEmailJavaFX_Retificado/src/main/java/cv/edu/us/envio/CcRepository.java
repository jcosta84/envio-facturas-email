package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public final class CcRepository {

    private final Database db;

    public CcRepository(Database db) {
        this.db = db;
    }

    public List<String> listar() throws SQLException {
        List<String> emails = new ArrayList<>();

        String sql =
            "SELECT email_cc FROM cc_email ORDER BY email_cc";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql);
             ResultSet rs = stmt.executeQuery()) {

            while (rs.next()) {
                emails.add(rs.getString("email_cc"));
            }
        }

        return emails;
    }

    public void inserir(String email) throws SQLException {
        String sql =
            "INSERT INTO cc_email(email_cc) VALUES(?)";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, email.trim());
            stmt.executeUpdate();
        }
    }

    public void eliminar(String email) throws SQLException {
        String sql =
            "DELETE FROM cc_email WHERE email_cc=?";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, email.trim());
            stmt.executeUpdate();
        }
    }
}
