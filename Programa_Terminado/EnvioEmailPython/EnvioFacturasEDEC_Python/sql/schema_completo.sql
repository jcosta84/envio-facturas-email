CREATE DATABASE IF NOT EXISTS factura_email CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE factura_email;
CREATE TABLE IF NOT EXISTS clientes(id INT AUTO_INCREMENT PRIMARY KEY,cil VARCHAR(50) NOT NULL UNIQUE,nome VARCHAR(100) NOT NULL,email VARCHAR(255) NOT NULL,arquivo_anexo VARCHAR(255));
CREATE TABLE IF NOT EXISTS cc_email(id INT AUTO_INCREMENT PRIMARY KEY,email_cc VARCHAR(255) NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS corpo_email(id INT PRIMARY KEY,conteudo TEXT);
CREATE TABLE IF NOT EXISTS relatorio(id INT AUTO_INCREMENT PRIMARY KEY,nome VARCHAR(150),email VARCHAR(255),cil VARCHAR(50),status VARCHAR(30),mensagem TEXT,data_envio DATETIME DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS usuarios(id INT AUTO_INCREMENT PRIMARY KEY,username VARCHAR(100) UNIQUE NOT NULL,password VARCHAR(255) NOT NULL,nome VARCHAR(150),email VARCHAR(150),nivel ENUM('admin','gerente') NOT NULL,estado ENUM('ativo','bloqueado','expirado','inativo') DEFAULT 'ativo',data_inicio DATE DEFAULT (CURRENT_DATE),data_fim DATE NULL,tentativas_login INT DEFAULT 0,ultimo_acesso DATETIME,ultimo_ip VARCHAR(45),criado_por INT NULL,data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP,data_alteracao DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS acessos(id INT AUTO_INCREMENT PRIMARY KEY,id_usuario INT NULL,username_informado VARCHAR(100),resultado VARCHAR(40),ip VARCHAR(50),sistema_operativo VARCHAR(100),navegador VARCHAR(100),observacao VARCHAR(255),data_hora DATETIME DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS logs(id INT AUTO_INCREMENT PRIMARY KEY,id_usuario INT NULL,operacao VARCHAR(100),modulo VARCHAR(100),descricao TEXT,ip VARCHAR(50),data_hora DATETIME DEFAULT CURRENT_TIMESTAMP);
INSERT INTO usuarios(username,password,nivel,estado) VALUES('admin','admin123','admin','ativo') ON DUPLICATE KEY UPDATE nivel='admin';
INSERT INTO corpo_email(id,conteudo) VALUES(1,'Anexamos o documento correspondente ao mês atual para que possa efetuar o pagamento.

Caso já tenha efetuado o pagamento, por favor, desconsidere esta mensagem.

Atenciosamente,
Direção Comercial - EDEC SUL') ON DUPLICATE KEY UPDATE id=id;
