package cv.edu.us.envio;

import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public final class ClienteRepository {
    private final Database db;

    public ClienteRepository(Database db) {
        this.db = db;
    }

    public boolean inserir(Cliente cliente) throws SQLException {
        String sql = "INSERT INTO clientes(cil, nome, email, arquivo_anexo) VALUES(?,?,?,?)";
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, cliente.cil());
            ps.setString(2, cliente.nome());
            ps.setString(3, cliente.email());
            ps.setString(4, cliente.arquivoAnexo());
            return ps.executeUpdate() > 0;
        }
    }

    public List<Cliente> listar() throws SQLException {
        List<Cliente> lista = new ArrayList<>();
        String sql = "SELECT cil, nome, email, arquivo_anexo FROM clientes ORDER BY nome";
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql);
             ResultSet rs = ps.executeQuery()) {
            while (rs.next()) {
                lista.add(new Cliente(
                        rs.getString("cil"),
                        rs.getString("nome"),
                        rs.getString("email"),
                        rs.getString("arquivo_anexo")
                ));
            }
        }
        return lista;
    }

    public boolean atualizarEmail(String cil, String email) throws SQLException {
        String sql = "UPDATE clientes SET email=? WHERE cil=?";
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, email.trim());
            ps.setString(2, cil);
            return ps.executeUpdate() > 0;
        }
    }

    public boolean eliminar(String cil) throws SQLException {
        String sql = "DELETE FROM clientes WHERE cil=?";
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, cil);
            return ps.executeUpdate() > 0;
        }
    }

    public boolean existeCil(String cil) throws SQLException {
        String sql = "SELECT 1 FROM clientes WHERE cil=? LIMIT 1";
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, cil);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next();
            }
        }
    }
}
