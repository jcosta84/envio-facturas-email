# Relatório do Sistema Envio Automático de E-mails 
## Documentação funcional, arquitetura, exemplos de código e estruturas de dados

### Curso: Engenharia Informática 
### Instituição: Universidade de Santiago 
### Unidade Curricular: Conceção e Análise de Algoritmos 
### Tema: Sistema de Gestão Comercial AQStore 

<br> <br> 

## Participantes 

- Paulo Costa Nº 7517 

## Docente 
 
 - Professor Valério Semedo 

<br><br>

### Praia, 2026

<br><br>

\newpage

# Índice

### - Introdução
### - Objetivos
### - Tecnologias e bibliotecas utilizadas
### - Estrutura geral do projeto
### - Configuração Maven
### - Ficheiro de configuração
### - Leitura do ficheiro INI
### - Ligação à base de dados
### - Modelo de cliente
### - Repositório de clientes
### - Gestão de destinatários CC
### - Gestão do corpo do e-mail
### - Serviço de envio de e-mails
### - Modelo de relatório
### - Repositório de relatórios
### - Exportação para Excel
### - Exportação para PDF
### - Interface JavaFX
### - Menu lateral
### - Cadastro de clientes
### - Consulta e operações CRUD
### - Tabela JavaFX
### - Edição do corpo do e-mail
### - Envio em lote
### - Estrutura de dados `ListaDupla`
### - Listas utilizadas no sistema
### - Comparação das listas
### - Validação de e-mail
### - Tratamento de exceções
### - Alertas
### - Classe `Launcher`
### - Arquitetura lógica
### - Fluxo de cadastro
### - Fluxo de envio
### - Segurança e boas práticas
### - Limitações identificadas
### - Melhorias futuras
### - Conclusão
### - Anexo A — Comandos de execução
### - Anexo B — Estrutura sugerida da base de dados
### - Anexo C — Conversão do Markdown para PDF

\newpage

# Introdução

O presente relatório documenta o desenvolvimento de uma aplicação desktop destinada à gestão de clientes e ao envio automático de documentos PDF por correio eletrónico. O sistema foi implementado em Java 21, utilizando JavaFX para a construção da interface gráfica, MySQL para armazenamento dos dados, Jakarta Mail para comunicação com o servidor SMTP, Apache POI para geração de ficheiros Excel e OpenPDF para criação de relatórios em PDF.

A aplicação foi organizada de forma modular, separando a interface gráfica, o acesso à base de dados, os serviços de envio e exportação, os modelos de dados e as estruturas de dados utilizadas no funcionamento interno. Esta organização melhora a legibilidade do código, facilita a manutenção e permite a evolução futura da solução.

O sistema permite cadastrar clientes, atualizar os respetivos endereços eletrónicos, eliminar registos, gerir destinatários em conhecimento, alterar o corpo padrão das mensagens, selecionar uma pasta com documentos PDF, efetuar o envio em lote e consultar o histórico de resultados.

\newpage

# Objetivos

## Objetivo geral

Desenvolver uma aplicação capaz de automatizar o envio de documentos PDF para clientes, associando cada ficheiro ao respetivo CIL e endereço de correio eletrónico.

## Objetivos específicos

- Desenvolver uma interface gráfica intuitiva;
- Efetuar a ligação a uma base de dados MySQL;
- Permitir o cadastro, consulta, atualização e eliminação de clientes;
- Gerir endereços de e-mail utilizados no campo CC;
- Permitir a edição do corpo padrão da mensagem;
- Validar os dados antes da sua utilização;
- Efetuar o envio de mensagens com documentos PDF em anexo;
- Registar os resultados dos envios;
- Filtrar relatórios por data;
- Exportar os resultados para Excel e PDF;
- Utilizar uma estrutura de dados própria do tipo lista duplamente ligada.

\newpage

# Tecnologias e bibliotecas utilizadas

| Tecnologia ou biblioteca | Aplicação no sistema |
|---|---|
| Java 21 | Linguagem principal do projeto |
| JavaFX 21.0.2 | Interface gráfica |
| Maven | Gestão das dependências e compilação |
| MySQL | Armazenamento persistente |
| JDBC | Comunicação entre Java e MySQL |
| Jakarta Mail | Envio de mensagens por SMTP |
| Apache POI | Exportação para Excel |
| OpenPDF | Exportação para PDF |
| Ficheiro INI | Configuração externa |
| Expressões regulares | Validação dos e-mails |

# Estrutura geral do projeto

```text
envio-email-javafx
│
├── pom.xml
├── config.ini
│
└── src/main/java/cv/edu/us/envio
    │
    ├── Launcher.java
    ├── MainApp.java
    ├── ConfigIni.java
    ├── Database.java
    ├── Cliente.java
    ├── RelatorioEnvio.java
    ├── ListaDupla.java
    ├── ClienteRepository.java
    ├── CcRepository.java
    ├── ConfiguracaoRepository.java
    ├── RelatorioRepository.java
    ├── EmailService.java
    └── ExportService.java
```

> **Nota explicativa:**  
> A estrutura do projeto separa responsabilidades. As classes `Cliente` e `RelatorioEnvio` representam dados, as classes terminadas em `Repository` tratam da base de dados, os serviços executam operações específicas e `MainApp` controla a interface JavaFX.

\newpage

# Configuração Maven

O ficheiro `pom.xml` define a versão do Java, as bibliotecas utilizadas e o plugin responsável pela execução do JavaFX.

```xml
<properties>
    <maven.compiler.release>21</maven.compiler.release>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
    <javafx.version>21.0.2</javafx.version>
</properties>
```

> **Nota explicativa:**  
> A propriedade `maven.compiler.release` garante que o projeto seja compilado para Java 21. A codificação UTF-8 permite utilizar corretamente acentos e caracteres especiais. A versão do JavaFX fica centralizada numa única propriedade, evitando repetições.

## Dependência do JavaFX

```xml
<dependency>
    <groupId>org.openjfx</groupId>
    <artifactId>javafx-controls</artifactId>
    <version>${javafx.version}</version>
</dependency>
```

> **Nota explicativa:**  
> A dependência `javafx-controls` disponibiliza componentes visuais como botões, campos de texto, tabelas, caixas de diálogo e seletores de data.

## Dependência do MySQL

```xml
<dependency>
    <groupId>com.mysql</groupId>
    <artifactId>mysql-connector-j</artifactId>
    <version>8.3.0</version>
</dependency>
```

> **Nota explicativa:**  
> O conector MySQL disponibiliza o driver JDBC necessário para que a aplicação Java consiga abrir ligações com o servidor MySQL.

## Dependência do Jakarta Mail

```xml
<dependency>
    <groupId>com.sun.mail</groupId>
    <artifactId>jakarta.mail</artifactId>
    <version>2.0.1</version>
</dependency>
```

> **Nota explicativa:**  
> Esta biblioteca permite criar mensagens MIME, autenticar no servidor SMTP, definir destinatários, adicionar anexos e enviar mensagens.

## Dependências de exportação

```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.2.5</version>
</dependency>

<dependency>
    <groupId>com.github.librepdf</groupId>
    <artifactId>openpdf</artifactId>
    <version>1.3.30</version>
</dependency>
```

> **Nota explicativa:**  
> O Apache POI é utilizado para gerar ficheiros `.xlsx`, enquanto o OpenPDF permite criar documentos PDF contendo títulos e tabelas.

\newpage

# Ficheiro de configuração

O sistema utiliza um ficheiro externo chamado `config.ini`.

```ini
[DATABASE]
HOST = endereco_do_servidor
PORT = 3306
DATABASE = nome_da_base
USERNAME = utilizador
PASSWORD = palavra_passe

[EMAIL]
REMETENTE = conta@email.com
SENHA_APP = palavra_passe_de_aplicacao
ASSUNTO = Documento mensal
SMTP_HOST = smtp.gmail.com
SMTP_PORT = 465
```

> **Nota explicativa:**  
> O ficheiro INI evita que os dados de configuração fiquem espalhados pelo código. Em ambientes reais, palavras-passe não devem ser publicadas no relatório nem incluídas num repositório Git público.

# Leitura do ficheiro INI

## Construtor da classe `ConfigIni`

```java
public ConfigIni(String caminho) {
    carregar(Path.of(caminho));
}
```

> **Nota explicativa:**  
> O construtor recebe o caminho do ficheiro e inicia imediatamente a leitura. A classe `Path` fornece uma forma moderna e segura de trabalhar com caminhos no Java.

## Verificação da existência do ficheiro

```java
if (!Files.exists(caminho)) {
    throw new IllegalStateException(
        "Ficheiro de configuração não encontrado: "
        + caminho.toAbsolutePath()
    );
}
```

> **Nota explicativa:**  
> Antes da leitura, o sistema verifica se o ficheiro existe. Caso contrário, é lançada uma exceção com o caminho completo, facilitando a identificação do problema.

## Leitura das linhas

```java
List<String> linhas =
    Files.readAllLines(
        caminho,
        StandardCharsets.UTF_8
    );
```

> **Nota explicativa:**  
> O método `readAllLines` carrega o conteúdo do ficheiro em memória. A codificação UTF-8 garante que caracteres portugueses sejam interpretados corretamente.

## Identificação das secções

```java
if (linha.startsWith("[") && linha.endsWith("]")) {
    secaoAtual = linha
        .substring(1, linha.length() - 1)
        .trim()
        .toUpperCase();

    secoes.putIfAbsent(
        secaoAtual,
        new HashMap<>()
    );
}
```

> **Nota explicativa:**  
> As linhas entre parênteses retos representam secções, por exemplo `[DATABASE]` e `[EMAIL]`. O nome é guardado em maiúsculas para tornar a consulta independente da combinação de letras.

## Separação entre chave e valor

```java
int separador = linha.indexOf('=');

String chave =
    linha.substring(0, separador)
         .trim()
         .toUpperCase();

String valor =
    linha.substring(separador + 1)
         .trim();
```

> **Nota explicativa:**  
> Cada linha de configuração é dividida no primeiro sinal de igualdade. A parte esquerda representa a chave e a parte direita representa o valor.

## Método de consulta

```java
public String get(
        String secao,
        String chave
) {
    Map<String, String> dados =
        secoes.get(
            secao.trim().toUpperCase()
        );

    if (dados == null) {
        return "";
    }

    return dados.getOrDefault(
        chave.trim().toUpperCase(),
        ""
    ).trim();
}
```

> **Nota explicativa:**  
> Este método permite consultar um valor através do nome da secção e da chave. Quando a secção ou a chave não existe, é devolvida uma string vazia, evitando referências nulas.

\newpage

# Ligação à base de dados

## Construção da URL JDBC

```java
this.url =
    "jdbc:mysql://" + host + ":" + port + "/" + database +
    "?useSSL=false" +
    "&serverTimezone=UTC" +
    "&allowPublicKeyRetrieval=true" +
    "&characterEncoding=UTF-8";
```

> **Nota explicativa:**  
> A URL JDBC contém o endereço do servidor, a porta e o nome da base de dados. Os parâmetros adicionais configuram o SSL, o fuso horário, a recuperação da chave pública e a codificação UTF-8.

## Método de ligação

```java
public Connection conectar()
        throws SQLException {

    return DriverManager.getConnection(
        url,
        username,
        password
    );
}
```

> **Nota explicativa:**  
> O método abre uma nova ligação ao MySQL. A exceção `SQLException` é propagada para que a camada que chamou o método possa informar o utilizador.

## Teste da ligação

```java
public void testar()
        throws SQLException {

    try (Connection ignored = conectar()) {
        // Ligação estabelecida.
    }
}
```

> **Nota explicativa:**  
> O bloco `try-with-resources` fecha automaticamente a ligação. Se o método terminar sem exceção, significa que o acesso ao servidor foi estabelecido.

# Modelo de cliente

```java
public record Cliente(
        String cil,
        String nome,
        String email,
        String arquivoAnexo
) {
}
```

> **Nota explicativa:**  
> Um `record` é uma estrutura imutável adequada para transportar dados. O Java cria automaticamente o construtor, métodos de acesso, `equals`, `hashCode` e `toString`.

## Normalização dos dados

```java
public Cliente {
    cil = cil == null ? "" : cil.trim();
    nome = nome == null ? "" : nome.trim();
    email = email == null ? "" : email.trim();
    arquivoAnexo =
        arquivoAnexo == null
            ? ""
            : arquivoAnexo.trim();
}
```

> **Nota explicativa:**  
> O construtor compacto do `record` remove espaços e substitui valores nulos por strings vazias. Isso reduz o risco de `NullPointerException`.

\newpage

# Repositório de clientes

## Inserção

```java
String sql =
    "INSERT INTO clientes(" +
    "cil, nome, email, arquivo_anexo" +
    ") VALUES(?,?,?,?)";
```

> **Nota explicativa:**  
> O comando inclui quatro parâmetros. A utilização de `?` permite associar os valores através de `PreparedStatement`, reduzindo o risco de SQL Injection.

```java
stmt.setString(1, cliente.cil());
stmt.setString(2, cliente.nome());
stmt.setString(3, cliente.email());
stmt.setString(4, cliente.arquivoAnexo());

return stmt.executeUpdate() > 0;
```

> **Nota explicativa:**  
> Os valores são associados pela ordem em que os parâmetros aparecem. `executeUpdate()` devolve o número de linhas afetadas.

## Listagem

```java
String sql =
    "SELECT cil, nome, email, arquivo_anexo " +
    "FROM clientes ORDER BY nome";
```

> **Nota explicativa:**  
> A consulta devolve os clientes por ordem alfabética, facilitando a consulta na interface.

```java
while (rs.next()) {
    lista.add(
        new Cliente(
            rs.getString("cil"),
            rs.getString("nome"),
            rs.getString("email"),
            rs.getString("arquivo_anexo")
        )
    );
}
```

> **Nota explicativa:**  
> Cada linha do `ResultSet` é convertida num objeto `Cliente` e adicionada à `ListaDupla`.

## Atualização

```java
String sql =
    "UPDATE clientes SET email=? WHERE cil=?";
```

> **Nota explicativa:**  
> O CIL identifica o cliente e o novo e-mail substitui o valor anterior.

## Eliminação

```java
String sql =
    "DELETE FROM clientes WHERE cil=?";
```

> **Nota explicativa:**  
> A eliminação utiliza o CIL como critério. O parâmetro evita a concatenação direta de valores no SQL.

## Verificação de duplicação

```java
String sql =
    "SELECT 1 FROM clientes " +
    "WHERE cil=? LIMIT 1";
```

> **Nota explicativa:**  
> A consulta verifica apenas a existência do registo. `LIMIT 1` evita continuar a pesquisa depois do primeiro resultado.

\newpage

# Gestão de destinatários CC

## Listagem

```java
String sql =
    "SELECT email_cc " +
    "FROM cc_email " +
    "ORDER BY email_cc";
```

> **Nota explicativa:**  
> Os endereços CC são carregados por ordem alfabética e apresentados numa `ListView`.

## Inserção

```java
String sql =
    "INSERT INTO cc_email(email_cc) VALUES(?)";
```

> **Nota explicativa:**  
> O endereço é inserido numa tabela própria, permitindo alterar a lista sem modificar o código.

## Eliminação

```java
String sql =
    "DELETE FROM cc_email WHERE email_cc=?";
```

> **Nota explicativa:**  
> O endereço selecionado é removido através do seu valor.

# Gestão do corpo do e-mail

## Obtenção

```java
String sql =
    "SELECT conteudo " +
    "FROM corpo_email " +
    "WHERE id=1";
```

> **Nota explicativa:**  
> O sistema utiliza um registo fixo com `id=1` para guardar o texto padrão.

## Inserção ou atualização

```java
String sql = """
    INSERT INTO corpo_email(id, conteudo)
    VALUES(1, ?)
    ON DUPLICATE KEY UPDATE
        conteudo=VALUES(conteudo)
    """;
```

> **Nota explicativa:**  
> Este comando MySQL funciona como `upsert`: cria o registo se não existir e atualiza-o quando já existe.

\newpage

# Serviço de envio de e-mails

## Valores SMTP com padrão

```java
String smtpHost =
    valorOuPadrao(
        config.get("EMAIL", "SMTP_HOST"),
        "smtp.gmail.com"
    );

String smtpPort =
    valorOuPadrao(
        config.get("EMAIL", "SMTP_PORT"),
        "465"
    );
```

> **Nota explicativa:**  
> Caso o host ou a porta não tenham sido definidos, o sistema utiliza os valores padrão do Gmail.

## Propriedades SMTP

```java
Properties props = new Properties();
props.put("mail.smtp.host", smtpHost);
props.put("mail.smtp.port", smtpPort);
props.put("mail.smtp.auth", "true");
props.put("mail.smtp.ssl.enable", "true");
```

> **Nota explicativa:**  
> Estas propriedades ativam autenticação e SSL. A porta 465 é normalmente usada para SMTP sobre SSL.

## Autenticação

```java
Session session = Session.getInstance(
    props,
    new Authenticator() {
        @Override
        protected PasswordAuthentication
        getPasswordAuthentication() {

            return new PasswordAuthentication(
                remetente,
                senhaApp
            );
        }
    }
);
```

> **Nota explicativa:**  
> A sessão recebe as credenciais da conta remetente. Para contas Gmail, deve ser utilizada uma palavra-passe de aplicação.

## Remetente

```java
mensagem.setFrom(
    new InternetAddress(remetente)
);
```

> **Nota explicativa:**  
> Define a conta que aparece como remetente.

## Destinatário principal

```java
mensagem.setRecipients(
    Message.RecipientType.TO,
    InternetAddress.parse(
        cliente.email()
    )
);
```

> **Nota explicativa:**  
> O destinatário principal é obtido do objeto `Cliente`.

## Destinatários CC

```java
for (String emailCc : emailsCc) {
    if (emailCc != null
            && !emailCc.isBlank()) {

        mensagem.addRecipient(
            Message.RecipientType.CC,
            new InternetAddress(
                emailCc.trim()
            )
        );
    }
}
```

> **Nota explicativa:**  
> Todos os endereços guardados na lista CC são adicionados à mensagem. Valores vazios são ignorados.

## Assunto

```java
mensagem.setSubject(
    assunto,
    StandardCharsets.UTF_8.name()
);
```

> **Nota explicativa:**  
> O assunto é codificado em UTF-8 para preservar os caracteres acentuados.

## Corpo personalizado

```java
texto.setText(
    "Olá " + cliente.nome() + ",\n\n" +
    "CIL: " + cliente.cil() + "\n\n" +
    corpoPadrao,
    StandardCharsets.UTF_8.name()
);
```

> **Nota explicativa:**  
> A mensagem inclui automaticamente o nome, o CIL e o conteúdo padrão guardado na base de dados.

## Anexo

```java
MimeBodyPart ficheiro =
    new MimeBodyPart();

ficiro.attachFile(anexo);
```

> **Nota explicativa:**  
> O ficheiro PDF é adicionado como parte MIME. No código real, a variável utilizada é `ficheiro`; qualquer escrita diferente deve ser corrigida.

## Mensagem multiparte

```java
Multipart multipart =
    new MimeMultipart();

multipart.addBodyPart(texto);
multipart.addBodyPart(ficheiro);

mensagem.setContent(multipart);
```

> **Nota explicativa:**  
> Uma mensagem multiparte pode conter simultaneamente texto e anexos.

## Envio

```java
Transport.send(mensagem);
```

> **Nota explicativa:**  
> Esta instrução inicia a comunicação com o servidor SMTP e envia a mensagem.

\newpage

# Modelo de relatório

```java
public record RelatorioEnvio(
        String nome,
        String email,
        String cil,
        String status,
        String mensagem,
        LocalDateTime dataEnvio
) {}
```

> **Nota explicativa:**  
> O `record` armazena todas as informações necessárias para apresentar e exportar o resultado de um envio.

# Repositório de relatórios

## Inserção

```java
String sql = """
    INSERT INTO relatorio(
        nome,
        email,
        cil,
        status,
        mensagem,
        data_envio
    )
    VALUES(?,?,?,?,?,?)
    """;
```

> **Nota explicativa:**  
> Cada resultado é gravado para permitir consulta posterior.

## Conversão da data

```java
stmt.setTimestamp(
    6,
    Timestamp.valueOf(
        relatorio.dataEnvio()
    )
);
```

> **Nota explicativa:**  
> O `LocalDateTime` é convertido em `Timestamp`, tipo compatível com JDBC.

## Construção dinâmica do filtro

```java
StringBuilder sql =
    new StringBuilder(
        "SELECT nome,email,cil,status," +
        "mensagem,data_envio " +
        "FROM relatorio WHERE 1=1"
    );
```

> **Nota explicativa:**  
> `WHERE 1=1` facilita a adição opcional de condições com `AND`.

## Filtro por data inicial

```java
if (inicio != null) {
    sql.append(" AND data_envio >= ?");
    parametros.add(
        Timestamp.valueOf(
            inicio.atStartOfDay()
        )
    );
}
```

> **Nota explicativa:**  
> A pesquisa inicia à meia-noite da data selecionada.

## Filtro por data final

```java
if (fim != null) {
    sql.append(" AND data_envio < ?");
    parametros.add(
        Timestamp.valueOf(
            fim.plusDays(1).atStartOfDay()
        )
    );
}
```

> **Nota explicativa:**  
> A utilização do início do dia seguinte inclui todos os horários da data final.

## Associação dos parâmetros

```java
for (int i = 0;
     i < parametros.size();
     i++) {

    stmt.setObject(
        i + 1,
        parametros.get(i)
    );
}
```

> **Nota explicativa:**  
> Os valores são associados na mesma ordem em que as condições foram acrescentadas.

\newpage

# Exportação para Excel

## Criação do livro

```java
try (Workbook workbook =
         new XSSFWorkbook()) {
}
```

> **Nota explicativa:**  
> `XSSFWorkbook` representa um ficheiro Excel no formato `.xlsx`.

## Criação da folha

```java
Sheet sheet =
    workbook.createSheet(
        "Relatório"
    );
```

> **Nota explicativa:**  
> O livro pode possuir várias folhas. Neste sistema é criada uma folha chamada “Relatório”.

## Cabeçalhos

```java
String[] cabecalhos = {
    "Data",
    "Nome",
    "Email",
    "CIL",
    "Status",
    "Mensagem"
};
```

> **Nota explicativa:**  
> O vetor define a ordem das colunas.

## Criação das células

```java
Row cabecalho =
    sheet.createRow(0);

for (int i = 0;
     i < cabecalhos.length;
     i++) {

    cabecalho
        .createCell(i)
        .setCellValue(
            cabecalhos[i]
        );
}
```

> **Nota explicativa:**  
> A primeira linha contém os títulos das colunas.

## Preenchimento dos dados

```java
for (RelatorioEnvio r : dados) {
    Row linha =
        sheet.createRow(numeroLinha++);

    linha.createCell(0)
         .setCellValue(
             r.dataEnvio()
              .format(FORMATO)
         );

    linha.createCell(1)
         .setCellValue(r.nome());
}
```

> **Nota explicativa:**  
> Para cada registo é criada uma nova linha.

## Ajuste automático

```java
for (int i = 0;
     i < cabecalhos.length;
     i++) {

    sheet.autoSizeColumn(i);
}
```

> **Nota explicativa:**  
> As larguras são ajustadas ao conteúdo.

## Gravação

```java
try (FileOutputStream out =
         new FileOutputStream(
             destino
         )) {

    workbook.write(out);
}
```

> **Nota explicativa:**  
> O livro é gravado no local selecionado pelo utilizador.

\newpage

# Exportação para PDF

## Criação

```java
Document documento =
    new Document();

PdfWriter.getInstance(
    documento,
    new FileOutputStream(destino)
);
```

> **Nota explicativa:**  
> O `Document` representa o PDF e o `PdfWriter` associa o documento ao ficheiro de destino.

## Abertura

```java
documento.open();
```

> **Nota explicativa:**  
> O documento deve ser aberto antes de receber conteúdo.

## Título

```java
documento.add(
    new Paragraph(
        "Relatório de Envios"
    )
);
```

> **Nota explicativa:**  
> É adicionado um título ao início do relatório.

## Tabela

```java
PdfPTable tabela =
    new PdfPTable(5);

tabela.setWidthPercentage(100);
```

> **Nota explicativa:**  
> A tabela possui cinco colunas e ocupa toda a largura disponível.

## Fecho

```java
documento.add(tabela);
documento.close();
```

> **Nota explicativa:**  
> O fecho finaliza a escrita e liberta os recursos.

\newpage

# Interface JavaFX

## Classe principal

```java
public class MainApp
        extends Application {
}
```

> **Nota explicativa:**  
> Toda aplicação JavaFX deve herdar de `Application`.

## Inicialização

```java
@Override
public void start(Stage stage) {
}
```

> **Nota explicativa:**  
> O JavaFX chama este método depois de inicializar a plataforma gráfica.

## Carregamento dos serviços

```java
config =
    new ConfigIni("config.ini");

database =
    new Database(config);

database.testar();
```

> **Nota explicativa:**  
> O sistema carrega primeiro a configuração e testa a base de dados antes de apresentar a interface.

## Instanciação dos repositórios

```java
clienteRepository =
    new ClienteRepository(database);

ccRepository =
    new CcRepository(database);

configuracaoRepository =
    new ConfiguracaoRepository(
        database
    );

relatorioRepository =
    new RelatorioRepository(
        database
    );
```

> **Nota explicativa:**  
> Todos os repositórios recebem a mesma instância de `Database`.

## Janela

```java
Scene scene =
    new Scene(
        root,
        1100,
        720
    );

stage.setTitle(
    "Gestão de Clientes " +
    "e Envio de E-mails"
);

stage.setScene(scene);
stage.show();
```

> **Nota explicativa:**  
> O `Stage` representa a janela e a `Scene` representa o seu conteúdo.

# Menu lateral

```java
Button cadastro =
    criarBotaoMenu(
        "Cadastro",
        () -> root.setCenter(
            criarCadastro()
        )
    );
```

> **Nota explicativa:**  
> Cada botão substitui o conteúdo central do `BorderPane`.

# Cadastro de clientes

## Campos

```java
TextField txtCil =
    new TextField();

TextField txtNome =
    new TextField();

TextField txtEmail =
    new TextField();
```

> **Nota explicativa:**  
> Estes campos recolhem os dados principais do cliente.

## Nome automático do PDF

```java
txtCil.textProperty()
    .addListener(
        (obs, antigo, novo) ->
            lblAnexo.setText(
                novo.isBlank()
                    ? ""
                    : novo.trim() + ".pdf"
            )
    );
```

> **Nota explicativa:**  
> O nome do anexo é atualizado automaticamente com base no CIL.

## Validação dos campos

```java
if (cil.isBlank()
        || nome.isBlank()
        || email.isBlank()) {

    mostrarAviso(
        "Preencha todos os campos."
    );

    return;
}
```

> **Nota explicativa:**  
> O fluxo é interrompido quando existe um campo obrigatório vazio.

## Verificação do e-mail

```java
if (!emailValido(email)) {
    mostrarAviso(
        "E-mail inválido."
    );
    return;
}
```

> **Nota explicativa:**  
> O sistema valida o formato antes de guardar.

## Criação do cliente

```java
Cliente cliente =
    new Cliente(
        cil,
        nome,
        email,
        cil + ".pdf"
    );
```

> **Nota explicativa:**  
> O ficheiro associado segue o padrão `CIL.pdf`.

\newpage

# Consulta e operações CRUD

## Filtro

```java
ListaDupla<Cliente> filtrados =
    clienteRepository
        .listar()
        .filtrar(
            cliente ->
                cliente.cil()
                       .toLowerCase()
                       .contains(termo)
                ||
                cliente.nome()
                       .toLowerCase()
                       .contains(termo)
        );
```

> **Nota explicativa:**  
> A pesquisa verifica o CIL e o nome, sem distinguir maiúsculas e minúsculas.

## Atualização

```java
clienteRepository
    .atualizarEmail(
        cliente.cil(),
        novoEmail
    );
```

> **Nota explicativa:**  
> O novo e-mail é guardado através do repositório.

## Eliminação

```java
clienteRepository
    .eliminar(
        cliente.cil()
    );
```

> **Nota explicativa:**  
> O cliente selecionado é identificado pelo CIL.

# Tabela JavaFX

```java
TableColumn<Cliente, String>
    colCil =
        new TableColumn<>("CIL");
```

> **Nota explicativa:**  
> Cada `TableColumn` representa uma propriedade exibida.

```java
colCil.setCellValueFactory(
    dados ->
        new SimpleStringProperty(
            dados.getValue().cil()
        )
);
```

> **Nota explicativa:**  
> A fábrica de células informa à tabela como obter o valor.

# Edição do corpo do e-mail

```java
TextArea txtCorpo =
    new TextArea();
```

> **Nota explicativa:**  
> O `TextArea` é adequado para textos com várias linhas.

```java
configuracaoRepository
    .guardarCorpoEmail(
        corpo
    );
```

> **Nota explicativa:**  
> O conteúdo é persistido para ser utilizado nos próximos envios.

\newpage

# Envio em lote

## Seleção da pasta

```java
DirectoryChooser chooser =
    new DirectoryChooser();

chooser.setTitle(
    "Selecionar pasta " +
    "dos ficheiros PDF"
);
```

> **Nota explicativa:**  
> O utilizador seleciona a pasta que contém os documentos.

## Filtro de PDFs

```java
File[] ficheiros =
    diretorio.listFiles(
        (dir, nome) ->
            nome.toLowerCase()
                .endsWith(".pdf")
    );
```

> **Nota explicativa:**  
> Apenas ficheiros com extensão `.pdf` são considerados.

## Associação pelo nome

```java
ficheirosPdf.put(
    ficheiro.getName(),
    ficheiro
);
```

> **Nota explicativa:**  
> Um mapa associa o nome do ficheiro ao objeto `File`, permitindo pesquisa rápida.

## Thread separada

```java
Thread threadEnvio =
    new Thread(() -> {
        // processamento
    });

threadEnvio.setDaemon(true);
threadEnvio.start();
```

> **Nota explicativa:**  
> O envio é executado fora da thread gráfica para evitar o bloqueio da interface.

## Leitura dos dados

```java
ListaDupla<Cliente>
    listaClientes =
        clienteRepository.listar();

ListaDupla<String>
    emailsCc =
        ccRepository.listar();

String corpoEmail =
    configuracaoRepository
        .obterCorpoEmail();
```

> **Nota explicativa:**  
> Antes do ciclo, são carregados os clientes, destinatários CC e o texto da mensagem.

## Correspondência do PDF

```java
File pdf =
    ficheirosPdf.get(
        cliente.arquivoAnexo()
    );
```

> **Nota explicativa:**  
> O nome guardado no cliente é utilizado para procurar o respetivo PDF.

## Estado ignorado

```java
resultado =
    new RelatorioEnvio(
        cliente.nome(),
        cliente.email(),
        cliente.cil(),
        "IGNORADO",
        "Anexo não encontrado: "
            + cliente.arquivoAnexo(),
        LocalDateTime.now()
    );
```

> **Nota explicativa:**  
> Quando não existe anexo, o cliente não é enviado e o motivo fica registado.

## Estado de sucesso

```java
resultado =
    new RelatorioEnvio(
        cliente.nome(),
        cliente.email(),
        cliente.cil(),
        "SUCESSO",
        "E-mail enviado " +
        "com sucesso.",
        LocalDateTime.now()
    );
```

> **Nota explicativa:**  
> Um registo de sucesso é criado depois de `EmailService.enviar()` terminar sem exceção.

## Estado de erro

```java
resultado =
    new RelatorioEnvio(
        cliente.nome(),
        cliente.email(),
        cliente.cil(),
        "ERRO",
        mensagemErro,
        LocalDateTime.now()
    );
```

> **Nota explicativa:**  
> A mensagem da exceção é guardada para facilitar o diagnóstico.

## Atualização segura da interface

```java
Platform.runLater(() -> {
    txtEstado.appendText(
        "[" + resultadoFinal.status()
        + "] "
        + resultadoFinal.nome()
        + " — "
        + resultadoFinal.mensagem()
        + System.lineSeparator()
    );
});
```

> **Nota explicativa:**  
> Componentes JavaFX só devem ser alterados na thread da aplicação. `Platform.runLater` agenda essa atualização.

\newpage

# Estrutura de dados `ListaDupla`

A aplicação utiliza uma implementação própria de uma lista duplamente ligada genérica.

```java
public class ListaDupla<T>
        extends AbstractList<T> {
}
```

> **Nota explicativa:**  
> O tipo genérico `T` permite utilizar a mesma estrutura para clientes, relatórios, strings e parâmetros SQL. A extensão de `AbstractList` mantém compatibilidade com APIs que esperam uma lista Java.

## Estrutura do nó

```java
private static final class No<T> {
    private T valor;
    private No<T> anterior;
    private No<T> proximo;

    private No(T valor) {
        this.valor = valor;
    }
}
```

> **Nota explicativa:**  
> Cada nó guarda o valor e referências para os dois vizinhos.

```text
NULL <- [Anterior | Valor | Próximo]
       <-> [Anterior | Valor | Próximo]
       <-> [Anterior | Valor | Próximo] -> NULL
```

## Referências principais

```java
private No<T> inicio;
private No<T> fim;
private int tamanho;
```

> **Nota explicativa:**  
> `inicio` aponta para o primeiro nó, `fim` para o último e `tamanho` guarda a quantidade de elementos.

## Verificação de vazio

```java
public boolean estaVazia() {
    return tamanho == 0;
}
```

> **Nota explicativa:**  
> O método evita comparar diretamente as referências externas.

## Inserção no início

```java
public void inserirInicio(T valor) {
    No<T> novo = new No<>(valor);

    if (inicio == null) {
        inicio = fim = novo;
    } else {
        novo.proximo = inicio;
        inicio.anterior = novo;
        inicio = novo;
    }

    tamanho++;
    modCount++;
}
```

> **Nota explicativa:**  
> A inserção no início possui complexidade `O(1)`, porque não percorre a lista.

## Inserção no fim

```java
@Override
public boolean add(T valor) {
    No<T> novo = new No<>(valor);

    if (fim == null) {
        inicio = fim = novo;
    } else {
        fim.proximo = novo;
        novo.anterior = fim;
        fim = novo;
    }

    tamanho++;
    modCount++;
    return true;
}
```

> **Nota explicativa:**  
> A referência `fim` permite inserir diretamente no último ponto da estrutura, com complexidade `O(1)`.

## Inserção por índice

```java
@Override
public void add(
        int indice,
        T valor
) {
    validarIndiceInsercao(indice);

    if (indice == tamanho) {
        add(valor);
        return;
    }

    if (indice == 0) {
        inserirInicio(valor);
        return;
    }
}
```

> **Nota explicativa:**  
> Os casos especiais de início e fim são tratados sem percurso.

## Acesso

```java
@Override
public T get(int indice) {
    validarIndice(indice);
    return obterNo(indice).valor;
}
```

> **Nota explicativa:**  
> O método obtém o nó correspondente e devolve o valor.

## Pesquisa otimizada do nó

```java
private No<T> obterNo(
        int indice
) {
    if (indice < tamanho / 2) {
        No<T> atual = inicio;

        for (int i = 0;
             i < indice;
             i++) {

            atual = atual.proximo;
        }

        return atual;
    }

    No<T> atual = fim;

    for (int i = tamanho - 1;
         i > indice;
         i--) {

        atual = atual.anterior;
    }

    return atual;
}
```

> **Nota explicativa:**  
> O percurso começa pelo lado mais próximo do índice. No pior caso, a complexidade é `O(n)`, mas são percorridos no máximo cerca de metade dos nós.

## Alteração

```java
@Override
public T set(
        int indice,
        T valor
) {
    validarIndice(indice);

    No<T> no =
        obterNo(indice);

    T antigo = no.valor;
    no.valor = valor;

    return antigo;
}
```

> **Nota explicativa:**  
> O método substitui o valor e devolve o anterior, conforme o contrato de `List`.

## Remoção por índice

```java
@Override
public T remove(int indice) {
    validarIndice(indice);
    No<T> no = obterNo(indice);
    return desligar(no);
}
```

> **Nota explicativa:**  
> Primeiro é localizado o nó e depois ele é desligado dos vizinhos.

## Remoção por valor

```java
@Override
public boolean remove(
        Object valor
) {
    No<T> atual = inicio;

    while (atual != null) {
        if (Objects.equals(
                atual.valor,
                valor
        )) {
            desligar(atual);
            return true;
        }

        atual = atual.proximo;
    }

    return false;
}
```

> **Nota explicativa:**  
> A lista é percorrida até encontrar o primeiro valor igual.

## Remoção no início e fim

```java
public T removerInicio() {
    if (inicio == null) {
        throw new NoSuchElementException(
            "A lista está vazia."
        );
    }

    return desligar(inicio);
}

public T removerFim() {
    if (fim == null) {
        throw new NoSuchElementException(
            "A lista está vazia."
        );
    }

    return desligar(fim);
}
```

> **Nota explicativa:**  
> Como as referências já existem, ambas as operações têm complexidade `O(1)`.

## Desligação do nó

```java
private T desligar(No<T> no) {
    No<T> anterior = no.anterior;
    No<T> proximo = no.proximo;

    if (anterior == null) {
        inicio = proximo;
    } else {
        anterior.proximo = proximo;
    }

    if (proximo == null) {
        fim = anterior;
    } else {
        proximo.anterior = anterior;
    }

    T valor = no.valor;

    no.valor = null;
    no.anterior = null;
    no.proximo = null;

    tamanho--;
    modCount++;

    return valor;
}
```

> **Nota explicativa:**  
> Este método centraliza a lógica de remoção. Ele funciona para nós no início, meio e fim.

## Pesquisa com `Predicate`

```java
public T pesquisar(
        Predicate<T> criterio
) {
    Objects.requireNonNull(
        criterio,
        "O critério não pode ser nulo."
    );

    for (T valor : this) {
        if (criterio.test(valor)) {
            return valor;
        }
    }

    return null;
}
```

> **Nota explicativa:**  
> O `Predicate` permite definir condições diferentes sem criar vários métodos de pesquisa.

## Filtragem

```java
public ListaDupla<T> filtrar(
        Predicate<T> criterio
) {
    ListaDupla<T> resultado =
        new ListaDupla<>();

    for (T valor : this) {
        if (criterio.test(valor)) {
            resultado.add(valor);
        }
    }

    return resultado;
}
```

> **Nota explicativa:**  
> A filtragem não altera a lista original. Uma nova lista é criada apenas com os elementos aprovados.

## Limpeza

```java
@Override
public void clear() {
    No<T> atual = inicio;

    while (atual != null) {
        No<T> proximo =
            atual.proximo;

        atual.anterior = null;
        atual.proximo = null;
        atual.valor = null;

        atual = proximo;
    }

    inicio = null;
    fim = null;
    tamanho = 0;
    modCount++;
}
```

> **Nota explicativa:**  
> As referências são removidas para permitir a libertação de memória pelo coletor de lixo.

## Iterador

```java
@Override
public Iterator<T> iterator() {
    return new Iterator<>() {
        private No<T> atual = inicio;

        @Override
        public boolean hasNext() {
            return atual != null;
        }

        @Override
        public T next() {
            if (atual == null) {
                throw new NoSuchElementException();
            }

            T valor = atual.valor;
            atual = atual.proximo;
            return valor;
        }
    };
}
```

> **Nota explicativa:**  
> O iterador permite utilizar a estrutura num ciclo `for-each`.

# Listas utilizadas no sistema

## Lista de clientes

```java
ListaDupla<Cliente> listaClientes =
    clienteRepository.listar();
```

> **Nota explicativa:**  
> Armazena os clientes carregados da base de dados e é utilizada no envio em lote.

## Lista de destinatários CC

```java
ListaDupla<String> emailsCc =
    ccRepository.listar();
```

> **Nota explicativa:**  
> Guarda endereços de e-mail em formato textual.

## Lista de resultados

```java
ListaDupla<RelatorioEnvio> resultados =
    new ListaDupla<>();
```

> **Nota explicativa:**  
> Regista temporariamente os resultados produzidos durante o envio.

## Lista de parâmetros SQL

```java
ListaDupla<Object> parametros =
    new ListaDupla<>();
```

> **Nota explicativa:**  
> Como os filtros podem utilizar tipos diferentes, a lista usa `Object`.

## Lista de linhas do ficheiro INI

```java
List<String> linhas =
    Files.readAllLines(
        caminho,
        StandardCharsets.UTF_8
    );
```

> **Nota explicativa:**  
> Esta é uma lista padrão da biblioteca Java, usada apenas durante a leitura do ficheiro.

## Lista de colunas JavaFX

```java
tabela.getColumns().addAll(
    colCil,
    colNome,
    colEmail,
    colAnexo
);
```

> **Nota explicativa:**  
> O componente `TableView` mantém internamente uma lista observável de colunas.

# Comparação das listas

| Lista | Tipo | Conteúdo | Finalidade |
|---|---|---|---|
| `ListaDupla<Cliente>` | Personalizada | Clientes | Consulta e envio |
| `ListaDupla<String>` | Personalizada | Endereços CC | Destinatários |
| `ListaDupla<RelatorioEnvio>` | Personalizada | Resultados | Relatórios |
| `ListaDupla<Object>` | Personalizada | Parâmetros | Consultas SQL |
| `List<String>` | Biblioteca Java | Linhas do INI | Configuração |
| `ObservableList` | JavaFX | Dados visuais | Tabelas e listas |

\newpage

# Validação de e-mail

```java
private boolean emailValido(
        String email
) {
    return Pattern.matches(
        "^[A-Za-z0-9+_.-]+@" +
        "[A-Za-z0-9.-]+$",
        email
    );
}
```

> **Nota explicativa:**  
> A expressão regular verifica uma estrutura básica de endereço eletrónico. Embora não cubra todos os casos do padrão oficial, é adequada para impedir muitos erros de digitação.

# Tratamento de exceções

```java
try {
    clienteRepository.inserir(
        cliente
    );
} catch (Exception ex) {
    mostrarErro(
        "Erro ao cadastrar",
        ex
    );
}
```

> **Nota explicativa:**  
> As exceções são capturadas na camada de interface para apresentar mensagens compreensíveis.

# Alertas

```java
private void mostrarAviso(
        String mensagem
) {
    new Alert(
        Alert.AlertType.WARNING,
        mensagem
    ).showAndWait();
}
```

> **Nota explicativa:**  
> Os alertas informam o utilizador sem encerrar a aplicação.

# Classe `Launcher`

```java
public class Launcher {
    public static void main(
            String[] args
    ) {
        MainApp.main(args);
    }
}
```

> **Nota explicativa:**  
> A classe auxiliar facilita a execução da aplicação em determinados ambientes de empacotamento.

\newpage

# Arquitetura lógica

```text
+------------------------------------------+
|                 MainApp                  |
| Interface, eventos e navegação JavaFX    |
+---------------------+--------------------+
                      |
       +--------------+--------------+
       |              |              |
       v              v              v
+-------------+ +-------------+ +-------------+
| Repositórios| | EmailService| |ExportService|
+------+------+ +------+------+ +------+------+
       |               |               |
       v               v               v
+-------------+ +-------------+ +-------------+
|    MySQL    | | Servidor    | |Excel / PDF  |
|             | | SMTP        | |             |
+-------------+ +-------------+ +-------------+
```

# Fluxo de cadastro

```text
Preencher formulário
        |
        v
Validar campos
        |
        v
Validar e-mail
        |
        v
Verificar CIL
        |
        v
Criar Cliente
        |
        v
Inserir no MySQL
        |
        v
Mostrar sucesso
```

# Fluxo de envio

```text
Selecionar pasta
        |
        v
Carregar PDFs
        |
        v
Carregar clientes
        |
        v
Carregar CC e corpo
        |
        v
Percorrer clientes
        |
        +---- PDF inexistente ----> IGNORADO
        |
        +---- Envio concluído ----> SUCESSO
        |
        +---- Exceção ------------> ERRO
        |
        v
Guardar relatório
        |
        v
Atualizar interface
```

# Segurança e boas práticas

## Prepared Statements

Os comandos SQL utilizam `PreparedStatement`, evitando a concatenação direta de entradas do utilizador.

## Fecho automático de recursos

As ligações, instruções, resultados e ficheiros utilizam `try-with-resources`.

## Configuração externa

Os dados variáveis ficam no `config.ini`. Contudo, recomenda-se proteger esse ficheiro e não o publicar.

## Validação

O sistema valida campos vazios, formato do e-mail, existência de clientes e presença de anexos.

## Processamento assíncrono

O envio ocorre numa thread separada, impedindo que a janela fique bloqueada.

# Limitações identificadas

- A palavra-passe de aplicação encontra-se em texto simples no INI;
- Não existe autenticação de utilizadores;
- O filtro de e-mail é básico;
- O envio é sequencial;
- Não existe barra de progresso;
- A eliminação não apresenta confirmação;
- Não existe controlo explícito de duplicação na tabela CC;
- O conteúdo do relatório PDF não inclui a mensagem detalhada;
- O sistema depende da correspondência exata entre CIL e nome do PDF.

# Melhorias futuras

- Encriptar credenciais;
- Utilizar variáveis de ambiente;
- Criar autenticação e perfis;
- Adicionar confirmação de eliminação;
- Mostrar progresso do envio;
- Permitir repetição apenas dos envios com erro;
- Guardar tentativas e duração;
- Criar modelos de mensagem;
- Permitir múltiplos anexos;
- Adicionar testes unitários;
- Utilizar pool de ligações;
- Adicionar logs estruturados;
- Criar instalador da aplicação.

\newpage

# Conclusão

O desenvolvimento do sistema permitiu aplicar conceitos de programação orientada a objetos, estruturas de dados, bases de dados, interfaces gráficas, comunicação por rede e geração de documentos.

A arquitetura adotada separa as principais responsabilidades, tornando o projeto mais organizado. Os repositórios tratam da persistência, os serviços concentram operações específicas e a interface coordena a interação com o utilizador.

A implementação da `ListaDupla` demonstra conhecimentos de estruturas de dados e algoritmos. A estrutura foi utilizada com diferentes tipos de informação graças ao uso de genéricos.

O sistema cumpre o objetivo de automatizar o envio de documentos PDF, mantendo um histórico dos resultados e permitindo a exportação das informações para formatos adequados à consulta e ao arquivo.

\newpage

# Anexo A — Comandos de execução

## Compilar

```powershell
mvn clean compile
```

> **Nota explicativa:**  
> Remove resultados anteriores e compila o projeto.

## Executar

```powershell
mvn javafx:run
```

> **Nota explicativa:**  
> Inicia a aplicação através do plugin JavaFX Maven.

## Gerar pacote

```powershell
mvn clean package
```

> **Nota explicativa:**  
> Cria o ficheiro JAR dentro da pasta `target`.

# Anexo B — Estrutura sugerida da base de dados

```sql
CREATE DATABASE factura_email
CHARACTER SET utf8mb4
COLLATE utf8mb4_unicode_ci;
```

```sql
CREATE TABLE clientes (
    cil VARCHAR(50) PRIMARY KEY,
    nome VARCHAR(200) NOT NULL,
    email VARCHAR(200) NOT NULL,
    arquivo_anexo VARCHAR(255) NOT NULL
);
```

```sql
CREATE TABLE cc_email (
    email_cc VARCHAR(200) PRIMARY KEY
);
```

```sql
CREATE TABLE corpo_email (
    id INT PRIMARY KEY,
    conteudo TEXT NOT NULL
);
```

```sql
CREATE TABLE relatorio (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(200),
    email VARCHAR(200),
    cil VARCHAR(50),
    status VARCHAR(30),
    mensagem TEXT,
    data_envio DATETIME NOT NULL
);
```

> **Nota explicativa:**  
> Estas tabelas correspondem às operações executadas pelos repositórios.

# Anexo C — Conversão do Markdown para PDF

```powershell
pandoc Relatorio_Sistema_Envio_Email_Completo.md `
  -o Relatorio_Sistema_Envio_Email_Completo.pdf `
  --toc `
  --number-sections `
  --pdf-engine=xelatex
```
