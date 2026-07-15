package cv.edu.us.envio;

import java.sql.*;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public final class RelatorioRepository {
    private final Database db;

    public RelatorioRepository(Database db) {
        this.db = db;
    }

    public void inserir(RelatorioEnvio r) throws SQLException {
        String sql = """
            INSERT INTO relatorio(nome,email,cil,status,mensagem,data_envio)
            VALUES(?,?,?,?,?,?)
            """;
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql)) {
            ps.setString(1, r.nome());
            ps.setString(2, r.email());
            ps.setString(3, r.cil());
            ps.setString(4, r.status());
            ps.setString(5, r.mensagem());
            ps.setTimestamp(6, Timestamp.valueOf(r.dataEnvio()));
            ps.executeUpdate();
        }
    }

    public List<RelatorioEnvio> listar(LocalDate inicio, LocalDate fim) throws SQLException {
        StringBuilder sql = new StringBuilder(
            "SELECT nome,email,cil,status,mensagem,data_envio FROM relatorio WHERE 1=1"
        );
        List<Object> params = new ArrayList<>();

        if (inicio != null) {
            sql.append(" AND data_envio >= ?");
            params.add(Timestamp.valueOf(inicio.atStartOfDay()));
        }
        if (fim != null) {
            sql.append(" AND data_envio < ?");
            params.add(Timestamp.valueOf(fim.plusDays(1).atStartOfDay()));
        }
        sql.append(" ORDER BY data_envio DESC");

        List<RelatorioEnvio> lista = new ArrayList<>();
        try (Connection c = db.conectar();
             PreparedStatement ps = c.prepareStatement(sql.toString())) {
            for (int i = 0; i < params.size(); i++) ps.setObject(i + 1, params.get(i));
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    lista.add(new RelatorioEnvio(
                        rs.getString("nome"),
                        rs.getString("email"),
                        rs.getString("cil"),
                        rs.getString("status"),
                        rs.getString("mensagem"),
                        rs.getTimestamp("data_envio").toLocalDateTime()
                    ));
                }
            }
        }
        return lista;
    }
}
