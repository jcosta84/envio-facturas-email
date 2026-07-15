package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Timestamp;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public final class RelatorioRepository {

    private final Database db;

    public RelatorioRepository(Database db) {
        this.db = db;
    }

    public void inserir(
            RelatorioEnvio relatorio
    ) throws SQLException {

        String sql = """
            INSERT INTO relatorio(
                nome,
                email,
                cil,
                status,
                mensagem,
                data_envio
            )
            VALUES(?,?,?,?,?,?)
            """;

        try (Connection conn = db.conectar();
             PreparedStatement stmt = conn.prepareStatement(sql)) {

            stmt.setString(1, relatorio.nome());
            stmt.setString(2, relatorio.email());
            stmt.setString(3, relatorio.cil());
            stmt.setString(4, relatorio.status());
            stmt.setString(5, relatorio.mensagem());
            stmt.setTimestamp(
                6,
                Timestamp.valueOf(relatorio.dataEnvio())
            );

            stmt.executeUpdate();
        }
    }

    public List<RelatorioEnvio> listar(
            LocalDate inicio,
            LocalDate fim
    ) throws SQLException {

        StringBuilder sql = new StringBuilder(
            "SELECT nome,email,cil,status,mensagem,data_envio " +
            "FROM relatorio WHERE 1=1"
        );

        List<Object> parametros = new ArrayList<>();

        if (inicio != null) {
            sql.append(" AND data_envio >= ?");
            parametros.add(
                Timestamp.valueOf(inicio.atStartOfDay())
            );
        }

        if (fim != null) {
            sql.append(" AND data_envio < ?");
            parametros.add(
                Timestamp.valueOf(
                    fim.plusDays(1).atStartOfDay()
                )
            );
        }

        sql.append(" ORDER BY data_envio DESC");

        List<RelatorioEnvio> lista = new ArrayList<>();

        try (Connection conn = db.conectar();
             PreparedStatement stmt =
                 conn.prepareStatement(sql.toString())) {

            for (int i = 0; i < parametros.size(); i++) {
                stmt.setObject(i + 1, parametros.get(i));
            }

            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    lista.add(
                        new RelatorioEnvio(
                            rs.getString("nome"),
                            rs.getString("email"),
                            rs.getString("cil"),
                            rs.getString("status"),
                            rs.getString("mensagem"),
                            rs.getTimestamp("data_envio")
                              .toLocalDateTime()
                        )
                    );
                }
            }
        }

        return lista;
    }
}
