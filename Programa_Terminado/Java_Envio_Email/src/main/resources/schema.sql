CREATE DATABASE IF NOT EXISTS factura_email
CHARACTER SET utf8mb4
COLLATE utf8mb4_unicode_ci;

USE factura_email;

CREATE TABLE IF NOT EXISTS clientes (
    cil VARCHAR(50) PRIMARY KEY,
    nome VARCHAR(150) NOT NULL,
    email VARCHAR(180) NOT NULL,
    arquivo_anexo VARCHAR(255) NOT NULL
);

CREATE TABLE IF NOT EXISTS cc_email (
    id INT AUTO_INCREMENT PRIMARY KEY,
    email_cc VARCHAR(180) NOT NULL
);

CREATE TABLE IF NOT EXISTS corpo_email (
    id INT PRIMARY KEY,
    conteudo LONGTEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS relatorio (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(150),
    email VARCHAR(180),
    cil VARCHAR(50),
    status VARCHAR(30),
    mensagem TEXT,
    data_envio DATETIME NOT NULL
);

INSERT INTO corpo_email (id, conteudo)
VALUES (
    1,
    'Anexamos o documento correspondente ao mês atual para que possa efetuar o pagamento.\n\n'
    'Lembramos que o não pagamento poderá resultar na interrupção do fornecimento.\n\n'
    'Caso já tenha efetuado o pagamento, por favor, desconsidere esta mensagem.\n\n'
    'Atenciosamente,\nDireção Comercial - EDEC SUL'
)
ON DUPLICATE KEY UPDATE conteudo = VALUES(conteudo);
