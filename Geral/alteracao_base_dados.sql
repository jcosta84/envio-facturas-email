-- Execute este script na mesma base de dados usada no config.ini.

-- Colunas necessárias na tabela usuarios (execute apenas as que ainda não existem).
ALTER TABLE usuarios
    ADD COLUMN estado ENUM('ativo','bloqueado','expirado','inativo') DEFAULT 'ativo',
    ADD COLUMN data_inicio DATE DEFAULT (CURRENT_DATE),
    ADD COLUMN data_fim DATE NULL,
    ADD COLUMN tentativas_login INT DEFAULT 0,
    ADD COLUMN ultimo_ip VARCHAR(45) NULL,
    ADD COLUMN criado_por INT NULL,
    ADD COLUMN data_alteracao DATETIME DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP;

-- Tabela de tentativas de autenticação.
CREATE TABLE IF NOT EXISTS acessos (
    id INT AUTO_INCREMENT PRIMARY KEY,
    id_usuario INT NULL,
    username_informado VARCHAR(100),
    resultado ENUM(
        'SUCESSO',
        'SENHA_INCORRETA',
        'UTILIZADOR_INEXISTENTE',
        'BLOQUEADO',
        'EXPIRADO',
        'INATIVO',
        'FORA_DO_PERIODO'
    ) NOT NULL,
    ip VARCHAR(45),
    sistema_operativo VARCHAR(100),
    navegador VARCHAR(100),
    data_hora DATETIME DEFAULT CURRENT_TIMESTAMP,
    observacao VARCHAR(255),
    CONSTRAINT fk_acessos_usuario
        FOREIGN KEY (id_usuario) REFERENCES usuarios(id)
        ON DELETE SET NULL
);

-- Tabela de operações efetuadas dentro do sistema.
CREATE TABLE IF NOT EXISTS logs (
    id INT AUTO_INCREMENT PRIMARY KEY,
    id_usuario INT NULL,
    operacao VARCHAR(100) NOT NULL,
    modulo VARCHAR(100),
    descricao TEXT,
    ip VARCHAR(45),
    data_hora DATETIME DEFAULT CURRENT_TIMESTAMP,
    CONSTRAINT fk_logs_usuario
        FOREIGN KEY (id_usuario) REFERENCES usuarios(id)
        ON DELETE SET NULL
);

CREATE INDEX idx_acessos_usuario ON acessos(id_usuario);
CREATE INDEX idx_acessos_data ON acessos(data_hora);
CREATE INDEX idx_logs_usuario ON logs(id_usuario);
CREATE INDEX idx_logs_data ON logs(data_hora);

-- Consultas para confirmar os registos:
SELECT * FROM acessos ORDER BY data_hora DESC;
SELECT * FROM logs ORDER BY data_hora DESC;
