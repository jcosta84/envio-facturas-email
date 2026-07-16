package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public final class Database {

    private final String url;
    private final String username;
    private final String password;

    public Database(ConfigIni config) {
        String host = config.get("DATABASE", "HOST");
        String port = config.get("DATABASE", "PORT");
        String database = config.get("DATABASE", "DATABASE");

        this.username = config.get("DATABASE", "USERNAME");
        this.password = config.get("DATABASE", "PASSWORD");

        this.url =
            "jdbc:mysql://" + host + ":" + port + "/" + database +
            "?useSSL=false" +
            "&serverTimezone=UTC" +
            "&allowPublicKeyRetrieval=true" +
            "&characterEncoding=UTF-8";
    }

    public Connection conectar() throws SQLException {
        return DriverManager.getConnection(url, username, password);
    }

    public void testar() throws SQLException {
        try (Connection ignored = conectar()) {
            // A ligação foi estabelecida com sucesso.
        }
    }
}
