package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Date;
import java.time.LocalDate;

/**
 * Repositório de autenticação e gestão de utilizadores.
 */
public final class UsuarioRepository {

    private final Database db;
    private final AuditoriaRepository auditoriaRepository;

    public UsuarioRepository(Database db) {
        this.db = db;
        this.auditoriaRepository = new AuditoriaRepository(db);
    }

    public Usuario autenticar(
            String username,
            String password
    ) throws SQLException {

        String usernameLimpo = username == null ? "" : username.trim();

        String sql = """
            SELECT id, username, password, nivel, estado,
                   data_inicio, data_fim
            FROM usuarios
            WHERE username = ?
            LIMIT 1
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, usernameLimpo);

            try (ResultSet rs = stmt.executeQuery()) {
                if (!rs.next()) {
                    auditoriaRepository.registrarAcesso(
                        null, usernameLimpo, "UTILIZADOR_INEXISTENTE",
                        "O nome de utilizador informado não existe."
                    );
                    return null;
                }

                int id = rs.getInt("id");

                if (!rs.getString("password").equals(password)) {
                    incrementarTentativas(id);
                    auditoriaRepository.registrarAcesso(
                        id, usernameLimpo, "SENHA_INCORRETA",
                        "Palavra-passe incorreta."
                    );
                    return null;
                }

                String estado = rs.getString("estado");
                if (!"ativo".equalsIgnoreCase(estado)) {
                    auditoriaRepository.registrarAcesso(
                        id, usernameLimpo, estado.toUpperCase(),
                        "A conta encontra-se no estado: " + estado
                    );
                    return null;
                }

                LocalDate hoje = LocalDate.now();
                Date inicioSql = rs.getDate("data_inicio");
                Date fimSql = rs.getDate("data_fim");

                if (inicioSql != null && hoje.isBefore(inicioSql.toLocalDate())) {
                    auditoriaRepository.registrarAcesso(
                        id, usernameLimpo, "FORA_DO_PERIODO",
                        "A conta ainda não atingiu a data de início."
                    );
                    return null;
                }

                if (fimSql != null && hoje.isAfter(fimSql.toLocalDate())) {
                    marcarExpirado(id);
                    auditoriaRepository.registrarAcesso(
                        id, usernameLimpo, "EXPIRADO",
                        "O período de utilização da conta terminou."
                    );
                    return null;
                }

                Usuario usuario = mapear(rs);
                atualizarUltimoAcesso(id);
                auditoriaRepository.registrarAcesso(
                    id, usernameLimpo, "SUCESSO",
                    "Login realizado com sucesso."
                );
                auditoriaRepository.registrarLog(
                    id, "LOGIN", "Autenticação",
                    "O utilizador entrou no sistema."
                );
                return usuario;
            }
        }
    }

    private void incrementarTentativas(int id) throws SQLException {
        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(
                 "UPDATE usuarios SET tentativas_login=tentativas_login+1 WHERE id=?")) {
            stmt.setInt(1, id);
            stmt.executeUpdate();
        }
    }

    private void atualizarUltimoAcesso(int id) throws SQLException {
        String sql = "UPDATE usuarios SET ultimo_acesso=NOW(), ultimo_ip=?, tentativas_login=0 WHERE id=?";
        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {
            stmt.setString(1, AuditoriaRepository.obterIpLocal());
            stmt.setInt(2, id);
            stmt.executeUpdate();
        }
    }

    private void marcarExpirado(int id) throws SQLException {
        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(
                 "UPDATE usuarios SET estado='expirado' WHERE id=?")) {
            stmt.setInt(1, id);
            stmt.executeUpdate();
        }
    }

    public ListaDupla<Usuario> listar()
            throws SQLException {

        ListaDupla<Usuario> usuarios =
            new ListaDupla<>();

        String sql = """
            SELECT id, username, password, nivel
            FROM usuarios
            ORDER BY username
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql);
             ResultSet rs =
                 stmt.executeQuery()) {

            while (rs.next()) {
                usuarios.add(
                    mapear(rs)
                );
            }
        }

        return usuarios;
    }

    public Usuario obterPorId(int id)
            throws SQLException {

        String sql = """
            SELECT id, username, password, nivel
            FROM usuarios
            WHERE id = ?
            LIMIT 1
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql)) {

            stmt.setInt(1, id);

            try (ResultSet rs =
                     stmt.executeQuery()) {

                return rs.next()
                    ? mapear(rs)
                    : null;
            }
        }
    }

    public void inserir(
            String username,
            String password,
            String nivel
    ) throws SQLException {

        String sql = """
            INSERT INTO usuarios(
                username,
                password,
                nivel
            )
            VALUES(?,?,?)
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql)) {

            stmt.setString(
                1,
                username.trim()
            );

            stmt.setString(
                2,
                password
            );

            stmt.setString(
                3,
                normalizarNivel(nivel)
            );

            stmt.executeUpdate();
        }
    }

    public boolean atualizar(
            int id,
            String username,
            String password,
            String nivel
    ) throws SQLException {

        String sql = """
            UPDATE usuarios
            SET username = ?,
                password = ?,
                nivel = ?
            WHERE id = ?
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql)) {

            stmt.setString(
                1,
                username.trim()
            );

            stmt.setString(
                2,
                password
            );

            stmt.setString(
                3,
                normalizarNivel(nivel)
            );

            stmt.setInt(4, id);

            return stmt.executeUpdate() > 0;
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
             PreparedStatement stmt =
                 conn.prepareStatement(sql)) {

            stmt.setString(
                1,
                novaSenha
            );

            stmt.setInt(
                2,
                id
            );

            stmt.setString(
                3,
                senhaAtual
            );

            return stmt.executeUpdate() > 0;
        }
    }

    public int contarAdministradores()
            throws SQLException {

        String sql = """
            SELECT COUNT(*)
            FROM usuarios
            WHERE LOWER(nivel)
                IN ('admin', 'administrador')
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql);
             ResultSet rs =
                 stmt.executeQuery()) {

            return rs.next()
                ? rs.getInt(1)
                : 0;
        }
    }

    public boolean eliminar(int id)
            throws SQLException {

        String sql =
            "DELETE FROM usuarios WHERE id=?";

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql)) {

            stmt.setInt(1, id);

            return stmt.executeUpdate() > 0;
        }
    }

    private Usuario mapear(ResultSet rs)
            throws SQLException {

        return new Usuario(
            rs.getInt("id"),
            rs.getString("username"),
            rs.getString("password"),
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
            nivel.trim()
                 .toLowerCase();

        if (!"admin".equals(valor)
                && !"gerente".equals(valor)) {

            throw new IllegalArgumentException(
                "O nível deve ser admin ou gerente."
            );
        }

        return valor;
    }
}
