package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 * Autenticação e gestão dos utilizadores.
 *
 * A tabela esperada possui:
 * id, username, password e nivel.
 */
public final class UsuarioRepository {

    private final Database db;

    public UsuarioRepository(Database db) {
        this.db = db;
    }

    public Usuario autenticar(
            String username,
            String password
    ) throws SQLException {

        String sql = """
            SELECT id, username, nivel
            FROM usuarios
            WHERE username = ?
              AND password = ?
            LIMIT 1
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, username.trim());
            stmt.setString(2, password);

            try (ResultSet rs = stmt.executeQuery()) {
                if (rs.next()) {
                    return mapear(rs);
                }
            }
        }

        return null;
    }

    public ListaDupla<Usuario> listar()
            throws SQLException {

        ListaDupla<Usuario> usuarios =
            new ListaDupla<>();

        String sql = """
            SELECT id, username, nivel
            FROM usuarios
            ORDER BY username
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql);
             ResultSet rs = stmt.executeQuery()) {

            while (rs.next()) {
                usuarios.add(mapear(rs));
            }
        }

        return usuarios;
    }

    public void inserir(
            String username,
            String password,
            String nivel
    ) throws SQLException {

        String nivelNormalizado =
            normalizarNivel(nivel);

        String sql = """
            INSERT INTO usuarios(
                username,
                password,
                nivel
            )
            VALUES(?,?,?)
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, username.trim());
            stmt.setString(2, password);
            stmt.setString(3, nivelNormalizado);
            stmt.executeUpdate();
        }
    }

    public boolean alterarSenhaPropria(
            int id,
            String senhaAtual,
            String novaSenha
    ) throws SQLException {

        String sql = """
            UPDATE usuarios
            SET password = ?
            WHERE id = ?
              AND password = ?
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, novaSenha);
            stmt.setInt(2, id);
            stmt.setString(3, senhaAtual);

            return stmt.executeUpdate() > 0;
        }
    }

    public boolean eliminar(int id)
            throws SQLException {

        String sql =
            "DELETE FROM usuarios WHERE id=?";

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setInt(1, id);
            return stmt.executeUpdate() > 0;
        }
    }

    private Usuario mapear(ResultSet rs)
            throws SQLException {

        return new Usuario(
            rs.getInt("id"),
            rs.getString("username"),
            rs.getString("nivel")
        );
    }

    private String normalizarNivel(
            String nivel
    ) {
        if (nivel == null) {
            throw new IllegalArgumentException(
                "Nível de acesso inválido."
            );
        }

        String valor =
            nivel.trim().toLowerCase();

        if (!"admin".equals(valor)
                && !"gerente".equals(valor)) {

            throw new IllegalArgumentException(
                "O nível deve ser admin ou gerente."
            );
        }

        return valor;
    }
}
