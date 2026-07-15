# Envio de E-mails — JavaFX

Conversão do programa Python/CustomTkinter para Java 21 + JavaFX.

## Funcionalidades
- Cadastro e CRUD de clientes
- Gestão de e-mails CC
- Edição do corpo do e-mail
- Envio em lote com PDF por CIL
- Persistência MySQL
- Relatórios por intervalo de datas
- Exportação Excel e PDF
- Estrutura de dados própria: `ListaLigada<T>`

## Preparação
1. Execute `src/main/resources/schema.sql` no MySQL.
2. Edite `src/main/resources/config.properties`.
3. No terminal:
   ```powershell
   mvn clean javafx:run
   ```

## Requisitos
- JDK 21
- Maven 3.9+
- MySQL 8
