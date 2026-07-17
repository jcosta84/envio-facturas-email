import mysql.connector
from services.config_service import ConfigService

class Database:
    def __init__(self, config: ConfigService): self.config = config
    def conectar(self):
        return mysql.connector.connect(
            host=self.config.get("DATABASE","HOST"),
            port=self.config.getint("DATABASE","PORT"),
            database=self.config.get("DATABASE","DATABASE"),
            user=self.config.get("DATABASE","USERNAME"),
            password=self.config.get("DATABASE","PASSWORD"),
            charset="utf8mb4",
            autocommit=False,
        )
    def testar(self):
        conn=self.conectar(); conn.close()
