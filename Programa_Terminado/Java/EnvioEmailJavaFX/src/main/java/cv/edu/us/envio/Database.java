package cv.edu.us.envio;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public final class Database {
    private final String url;
    private final String user;
    private final String password;

    public Database(AppConfig config) {
        this.url = "jdbc:mysql://" + config.get("db.host") + ":" +
                config.get("db.port") + "/" + config.get("db.name") +
                "?useSSL=false&serverTimezone=UTC&allowPublicKeyRetrieval=true&characterEncoding=UTF-8";
        this.user = config.get("db.user");
        this.password = config.get("db.password");
    }

    public Connection conectar() throws SQLException {
        return DriverManager.getConnection(url, user, password);
    }

    public void testar() throws SQLException {
        try (Connection ignored = conectar()) {
            // Ligação validada.
        }
    }
}
