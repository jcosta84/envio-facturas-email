# Envio de Faturas por E-mail — JavaFX

Conversão do programa Python para Java 21 com interface JavaFX.

## Funcionalidades

- Cadastro de clientes;
- CRUD de clientes;
- Gestão de endereços em CC;
- Edição do corpo do e-mail;
- Envio em lote com anexos PDF;
- Persistência em MySQL;
- Relatórios por período;
- Exportação para Excel;
- Exportação para PDF;
- Estrutura própria `ListaDupla<T>`.

## Configuração

O ficheiro `config.ini` deve ficar na raiz do projeto, ao lado do `pom.xml`.

Exemplo:

```ini
[DATABASE]
HOST = 100.116.112.107
PORT = 3306
DATABASE = factura_email
USERNAME = ALTERAR
PASSWORD = ALTERAR

[EMAIL]
REMETENTE = seuemail@gmail.com
SENHA_APP = ALTERAR
ASSUNTO = Factura de Energia da Empresa EDEC
SMTP_HOST = smtp.gmail.com
SMTP_PORT = 465
```

## Preparação

1. Execute `src/main/resources/schema.sql` no MySQL.
2. Preencha os dados reais no `config.ini`.
3. Configure o `JAVA_HOME` para um JDK 21 ou superior.
4. Execute:

```powershell
mvn clean javafx:run
```

## Estrutura de dados

A classe `ListaDupla<T>` foi implementada manualmente com nós anteriores e seguintes.

Ela substitui o uso de `ArrayList` na camada de dados e oferece:
- inserção no início, fim e posição;
- remoção no início, fim e posição;
- pesquisa por critério;
- filtragem;
- percurso nos dois sentidos;
- compatibilidade com JavaFX por meio de `AbstractList`.
