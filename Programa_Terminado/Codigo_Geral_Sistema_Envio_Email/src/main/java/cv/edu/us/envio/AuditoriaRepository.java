package cv.edu.us.envio;

import java.net.InetAddress;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Types;

/** Regista tentativas de acesso e operações efetuadas no sistema. */
public final class AuditoriaRepository {

    private final Database db;

    public AuditoriaRepository(Database db) {
        this.db = db;
    }

    public void registrarAcesso(
            Integer idUsuario,
            String usernameInformado,
            String resultado,
            String observacao
    ) throws SQLException {

        String sql = """
            INSERT INTO acessos(
                id_usuario,
                username_informado,
                resultado,
                ip,
                sistema_operativo,
                navegador,
                observacao,
                data_hora
            )
            VALUES(?,?,?,?,?,?,?,NOW())
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            if (idUsuario == null) {
                stmt.setNull(1, Types.INTEGER);
            } else {
                stmt.setInt(1, idUsuario);
            }

            stmt.setString(2, limitar(usernameInformado, 100));
            stmt.setString(3, resultado);
            stmt.setString(4, obterIpLocal());
            stmt.setString(5, limitar(System.getProperty("os.name"), 100));
            stmt.setString(6, "Aplicação JavaFX");
            stmt.setString(7, limitar(observacao, 255));
            stmt.executeUpdate();
        }
    }

    public void registrarLog(
            Integer idUsuario,
            String operacao,
            String modulo,
            String descricao
    ) throws SQLException {

        String sql = """
            INSERT INTO logs(
                id_usuario,
                operacao,
                modulo,
                descricao,
                ip,
                data_hora
            )
            VALUES(?,?,?,?,?,NOW())
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            if (idUsuario == null) {
                stmt.setNull(1, Types.INTEGER);
            } else {
                stmt.setInt(1, idUsuario);
            }

            stmt.setString(2, limitar(operacao, 100));
            stmt.setString(3, limitar(modulo, 100));
            stmt.setString(4, descricao);
            stmt.setString(5, obterIpLocal());
            stmt.executeUpdate();
        }
    }

    public static String obterIpLocal() {
        try {
            return InetAddress.getLocalHost().getHostAddress();
        } catch (Exception e) {
            return "DESCONHECIDO";
        }
    }

    private static String limitar(String valor, int tamanho) {
        if (valor == null) return null;
        return valor.length() <= tamanho ? valor : valor.substring(0, tamanho);
    }
}
