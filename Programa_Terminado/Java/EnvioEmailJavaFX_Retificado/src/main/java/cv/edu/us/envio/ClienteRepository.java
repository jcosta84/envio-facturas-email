package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public final class ClienteRepository {

    private final Database db;

    public ClienteRepository(Database db) {
        this.db = db;
    }

    public boolean inserir(Cliente cliente) throws SQLException {
        String sql =
            "INSERT INTO clientes(cil, nome, email, arquivo_anexo) " +
            "VALUES(?,?,?,?)";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, cliente.cil());
            stmt.setString(2, cliente.nome());
            stmt.setString(3, cliente.email());
            stmt.setString(4, cliente.arquivoAnexo());

            return stmt.executeUpdate() > 0;
        }
    }

    public List<Cliente> listar() throws SQLException {
        List<Cliente> lista = new ArrayList<>();

        String sql =
            "SELECT cil, nome, email, arquivo_anexo " +
            "FROM clientes ORDER BY nome";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql);
             ResultSet rs = stmt.executeQuery()) {

            while (rs.next()) {
                lista.add(
                    new Cliente(
                        rs.getString("cil"),
                        rs.getString("nome"),
                        rs.getString("email"),
                        rs.getString("arquivo_anexo")
                    )
                );
            }
        }

        return lista;
    }

    public boolean atualizarEmail(
            String cil,
            String novoEmail
    ) throws SQLException {

        String sql =
            "UPDATE clientes SET email=? WHERE cil=?";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, novoEmail.trim());
            stmt.setString(2, cil.trim());

            return stmt.executeUpdate() > 0;
        }
    }

    public boolean eliminar(String cil) throws SQLException {
        String sql = "DELETE FROM clientes WHERE cil=?";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, cil.trim());
            return stmt.executeUpdate() > 0;
        }
    }

    public boolean existeCil(String cil) throws SQLException {
        String sql =
            "SELECT 1 FROM clientes WHERE cil=? LIMIT 1";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, cil.trim());

            try (ResultSet rs = stmt.executeQuery()) {
                return rs.next();
            }
        }
    }
}
