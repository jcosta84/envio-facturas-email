package cv.edu.us.envio;

import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public final class CcRepository {
    private final Database db;

    public CcRepository(Database db) {
        this.db = db;
    }

    public List<String> listar() throws SQLException {
        List<String> emails = new ArrayList<>();
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement("SELECT email_cc FROM cc_email ORDER BY email_cc");
             ResultSet rs = ps.executeQuery()) {
            while (rs.next()) emails.add(rs.getString("email_cc"));
        }
        return emails;
    }

    public void inserir(String email) throws SQLException {
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement("INSERT INTO cc_email(email_cc) VALUES(?)")) {
            ps.setString(1, email.trim());
            ps.executeUpdate();
        }
    }

    public void eliminar(String email) throws SQLException {
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement("DELETE FROM cc_email WHERE email_cc=?")) {
            ps.setString(1, email);
            ps.executeUpdate();
        }
    }
}
