CREATE TABLE IF NOT EXISTS usuarios (
    id INT AUTO_INCREMENT PRIMARY KEY,
    username VARCHAR(100) NOT NULL UNIQUE,
    password VARCHAR(255) NOT NULL,
    nivel ENUM('admin', 'gerente') NOT NULL DEFAULT 'gerente'
);

INSERT INTO usuarios(
    username,
    password,
    nivel
)
VALUES(
    'admin',
    'admin123',
    'admin'
)
ON DUPLICATE KEY UPDATE
    nivel = 'admin';

UPDATE usuarios
SET nivel = 'admin'
WHERE LOWER(nivel) = 'administrador';
