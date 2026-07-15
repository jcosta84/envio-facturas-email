package cv.edu.us.envio;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

public class MainApp extends Application {

    private final BorderPane root = new BorderPane();

    private ConfigIni config;
    private Database database;
    private ClienteRepository clienteRepository;
    private CcRepository ccRepository;
    private ConfiguracaoRepository configuracaoRepository;
    private RelatorioRepository relatorioRepository;
    private EmailService emailService;

    private final ExportService exportService =
        new ExportService();

    @Override
    public void start(Stage stage) {
        try {
            config = new ConfigIni("config.ini");

            database = new Database(config);
            database.testar();

            clienteRepository =
                new ClienteRepository(database);

            ccRepository =
                new CcRepository(database);

            configuracaoRepository =
                new ConfiguracaoRepository(database);

            relatorioRepository =
                new RelatorioRepository(database);

            emailService =
                new EmailService(config);

            root.setLeft(criarMenu());
            root.setCenter(criarCadastro());

            Scene scene =
                new Scene(root, 1100, 720);

            stage.setTitle(
                "Gestão de Clientes e Envio de E-mails"
            );

            stage.setScene(scene);
            stage.show();

        } catch (Exception e) {
            mostrarErro("Erro ao iniciar", e);
            Platform.exit();
        }
    }

    private Node criarMenu() {
        VBox menu = new VBox(10);
        menu.setPadding(new Insets(16));
        menu.setPrefWidth(220);

        Label titulo = new Label("MENU");
        titulo.setStyle(
            "-fx-font-size: 18px;" +
            "-fx-font-weight: bold;"
        );

        Button cadastro =
            criarBotaoMenu(
                "Cadastro",
                () -> root.setCenter(criarCadastro())
            );

        Button cc =
            criarBotaoMenu(
                "CC Conhecimento",
                () -> root.setCenter(criarCc())
            );

        Button consulta =
            criarBotaoMenu(
                "Consulta",
                () -> root.setCenter(criarConsulta())
            );

        Button corpoEmail =
            criarBotaoMenu(
                "Corpo do E-mail",
                () -> root.setCenter(criarCorpoEmail())
            );

        Button envio =
            criarBotaoMenu(
                "Enviar E-mails",
                () -> root.setCenter(criarEnvio())
            );

        Button relatorio =
            criarBotaoMenu(
                "Relatório",
                () -> root.setCenter(criarRelatorio())
            );

        menu.getChildren().addAll(
            titulo,
            cadastro,
            cc,
            consulta,
            corpoEmail,
            envio,
            relatorio
        );

        return menu;
    }

    private Button criarBotaoMenu(
            String texto,
            Runnable acao
    ) {
        Button botao = new Button(texto);
        botao.setMaxWidth(Double.MAX_VALUE);
        botao.setOnAction(e -> acao.run());
        return botao;
    }

    private Node criarCadastro() {
        VBox box = criarPagina("Cadastro de Cliente");

        TextField txtCil = new TextField();
        TextField txtNome = new TextField();
        TextField txtEmail = new TextField();

        Label lblAnexo = new Label();

        txtCil.textProperty().addListener(
            (obs, antigo, novo) ->
                lblAnexo.setText(
                    novo.isBlank()
                        ? ""
                        : novo.trim() + ".pdf"
                )
        );

        GridPane formulario = new GridPane();
        formulario.setHgap(10);
        formulario.setVgap(10);

        formulario.addRow(
            0,
            new Label("CIL:"),
            txtCil
        );

        formulario.addRow(
            1,
            new Label("Nome:"),
            txtNome
        );

        formulario.addRow(
            2,
            new Label("E-mail:"),
            txtEmail
        );

        formulario.addRow(
            3,
            new Label("Nome do PDF:"),
            lblAnexo
        );

        Button btnGuardar =
            new Button("Cadastrar");

        Button btnLimpar =
            new Button("Limpar");

        btnGuardar.setOnAction(e -> {
            String cil = txtCil.getText().trim();
            String nome = txtNome.getText().trim();
            String email = txtEmail.getText().trim();

            if (cil.isBlank() ||
                nome.isBlank() ||
                email.isBlank()) {

                mostrarAviso(
                    "Preencha todos os campos."
                );

                return;
            }

            if (!emailValido(email)) {
                mostrarAviso("E-mail inválido.");
                return;
            }

            try {
                if (clienteRepository.existeCil(cil)) {
                    mostrarAviso(
                        "Já existe um cliente com este CIL."
                    );
                    return;
                }

                Cliente cliente = new Cliente(
                    cil,
                    nome,
                    email,
                    cil + ".pdf"
                );

                clienteRepository.inserir(cliente);

                mostrarInfo(
                    "Cliente cadastrado com sucesso."
                );

                txtCil.clear();
                txtNome.clear();
                txtEmail.clear();

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao cadastrar",
                    ex
                );
            }
        });

        btnLimpar.setOnAction(e -> {
            txtCil.clear();
            txtNome.clear();
            txtEmail.clear();
        });

        HBox botoes =
            new HBox(10, btnGuardar, btnLimpar);

        box.getChildren().addAll(
            formulario,
            botoes
        );

        return box;
    }

    private Node criarCc() {
        VBox box =
            criarPagina(
                "E-mails para Conhecimento (CC)"
            );

        ListView<String> lista =
            new ListView<>();

        Runnable carregar = () -> {
            try {
                lista.setItems(
                    FXCollections.observableArrayList(
                        ccRepository.listar()
                    )
                );
            } catch (Exception e) {
                mostrarErro(
                    "Erro ao carregar CC",
                    e
                );
            }
        };

        TextField txtNovo =
            new TextField();

        txtNovo.setPromptText(
            "novo@email.com"
        );

        Button btnAdicionar =
            new Button("Adicionar");

        Button btnEliminar =
            new Button("Eliminar selecionado");

        btnAdicionar.setOnAction(e -> {
            String email =
                txtNovo.getText().trim();

            if (!emailValido(email)) {
                mostrarAviso("E-mail inválido.");
                return;
            }

            try {
                ccRepository.inserir(email);
                txtNovo.clear();
                carregar.run();

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao adicionar CC",
                    ex
                );
            }
        });

        btnEliminar.setOnAction(e -> {
            String email =
                lista.getSelectionModel()
                     .getSelectedItem();

            if (email == null) {
                mostrarAviso(
                    "Selecione um e-mail."
                );
                return;
            }

            try {
                ccRepository.eliminar(email);
                carregar.run();

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao eliminar CC",
                    ex
                );
            }
        });

        box.getChildren().addAll(
            lista,
            new HBox(
                10,
                txtNovo,
                btnAdicionar,
                btnEliminar
            )
        );

        VBox.setVgrow(lista, Priority.ALWAYS);

        carregar.run();
        return box;
    }

    private Node criarConsulta() {
        VBox box =
            criarPagina(
                "Consulta e CRUD de Clientes"
            );

        TextField txtFiltro =
            new TextField();

        txtFiltro.setPromptText(
            "Filtrar por CIL ou nome"
        );

        TableView<Cliente> tabela =
            criarTabelaClientes();

        Runnable carregar = () -> {
            try {
                tabela.setItems(
                    FXCollections.observableArrayList(
                        clienteRepository.listar()
                    )
                );
            } catch (Exception e) {
                mostrarErro(
                    "Erro ao carregar clientes",
                    e
                );
            }
        };

        txtFiltro.textProperty().addListener(
            (obs, antigo, valor) -> {
                try {
                    String termo =
                        valor.toLowerCase();

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

                    tabela.setItems(
                        FXCollections
                            .observableArrayList(
                                filtrados
                            )
                    );

                } catch (Exception e) {
                    mostrarErro(
                        "Erro ao filtrar",
                        e
                    );
                }
            }
        );

        TextField txtNovoEmail =
            new TextField();

        txtNovoEmail.setPromptText(
            "Novo e-mail"
        );

        Button btnAtualizar =
            new Button("Atualizar e-mail");

        Button btnEliminar =
            new Button("Eliminar");

        btnAtualizar.setOnAction(e -> {
            Cliente cliente =
                tabela.getSelectionModel()
                      .getSelectedItem();

            String novoEmail =
                txtNovoEmail.getText().trim();

            if (cliente == null ||
                !emailValido(novoEmail)) {

                mostrarAviso(
                    "Selecione um cliente e informe " +
                    "um e-mail válido."
                );

                return;
            }

            try {
                clienteRepository.atualizarEmail(
                    cliente.cil(),
                    novoEmail
                );

                carregar.run();

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao atualizar",
                    ex
                );
            }
        });

        btnEliminar.setOnAction(e -> {
            Cliente cliente =
                tabela.getSelectionModel()
                      .getSelectedItem();

            if (cliente == null) {
                mostrarAviso(
                    "Selecione um cliente."
                );
                return;
            }

            try {
                clienteRepository.eliminar(
                    cliente.cil()
                );

                carregar.run();

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao eliminar",
                    ex
                );
            }
        });

        box.getChildren().addAll(
            txtFiltro,
            tabela,
            new HBox(
                10,
                txtNovoEmail,
                btnAtualizar,
                btnEliminar
            )
        );

        VBox.setVgrow(
            tabela,
            Priority.ALWAYS
        );

        carregar.run();
        return box;
    }

    private TableView<Cliente>
    criarTabelaClientes() {

        TableView<Cliente> tabela =
            new TableView<>();

        TableColumn<Cliente, String> colCil =
            new TableColumn<>("CIL");

        colCil.setCellValueFactory(
            dados ->
                new SimpleStringProperty(
                    dados.getValue().cil()
                )
        );

        TableColumn<Cliente, String> colNome =
            new TableColumn<>("Nome");

        colNome.setCellValueFactory(
            dados ->
                new SimpleStringProperty(
                    dados.getValue().nome()
                )
        );

        TableColumn<Cliente, String> colEmail =
            new TableColumn<>("E-mail");

        colEmail.setCellValueFactory(
            dados ->
                new SimpleStringProperty(
                    dados.getValue().email()
                )
        );

        TableColumn<Cliente, String> colAnexo =
            new TableColumn<>("Anexo");

        colAnexo.setCellValueFactory(
            dados ->
                new SimpleStringProperty(
                    dados.getValue()
                         .arquivoAnexo()
                )
        );

        tabela.getColumns().addAll(
            colCil,
            colNome,
            colEmail,
            colAnexo
        );

        tabela.setColumnResizePolicy(
            TableView
                .CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS
        );

        return tabela;
    }

    private Node criarCorpoEmail() {
        VBox box =
            criarPagina("Corpo do E-mail");

        TextArea txtCorpo =
            new TextArea();

        Button btnRecarregar =
            new Button("Recarregar");

        Button btnGuardar =
            new Button("Guardar alterações");

        Runnable recarregar = () -> {
            try {
                txtCorpo.setText(
                    configuracaoRepository
                        .obterCorpoEmail()
                );

            } catch (Exception e) {
                mostrarErro(
                    "Erro ao carregar o corpo",
                    e
                );
            }
        };

        btnRecarregar.setOnAction(
            e -> recarregar.run()
        );

        btnGuardar.setOnAction(e -> {
            String corpo =
                txtCorpo.getText().trim();

            if (corpo.isBlank()) {
                mostrarAviso(
                    "O corpo do e-mail não pode " +
                    "ficar vazio."
                );
                return;
            }

            try {
                configuracaoRepository
                    .guardarCorpoEmail(corpo);

                mostrarInfo(
                    "Corpo do e-mail atualizado."
                );

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao guardar",
                    ex
                );
            }
        });

        box.getChildren().addAll(
            txtCorpo,
            new HBox(
                10,
                btnGuardar,
                btnRecarregar
            )
        );

        VBox.setVgrow(
            txtCorpo,
            Priority.ALWAYS
        );

        recarregar.run();
        return box;
    }

    private Node criarEnvio() {

        VBox box = criarPagina("Envio de E-mails com PDF");

        Label lblPasta = new Label("Nenhuma pasta selecionada.");

        TextArea txtEstado = new TextArea();
        txtEstado.setEditable(false);
        txtEstado.setWrapText(true);
        txtEstado.setPromptText("O resultado dos envios será apresentado aqui.");

        Map<String, File> ficheirosPdf = new HashMap<>();

        Button btnEscolher = new Button("Selecionar diretório");
        Button btnEnviar = new Button("Enviar em lote");

        btnEnviar.setDisable(true);

        /*
        * Selecionar a pasta que contém os ficheiros PDF.
        */
        btnEscolher.setOnAction(e -> {

            try {
                DirectoryChooser chooser = new DirectoryChooser();
                chooser.setTitle("Selecionar pasta dos ficheiros PDF");

                File diretorio = chooser.showDialog(
                        root.getScene().getWindow()
                );

                if (diretorio == null) {
                    return;
                }

                ficheirosPdf.clear();

                File[] ficheiros = diretorio.listFiles(
                        (dir, nome) ->
                                nome.toLowerCase().endsWith(".pdf")
                );

                if (ficheiros != null) {
                    for (File ficheiro : ficheiros) {
                        ficheirosPdf.put(
                                ficheiro.getName(),
                                ficheiro
                        );
                    }
                }

                lblPasta.setText(
                        diretorio.getAbsolutePath()
                                + " — "
                                + ficheirosPdf.size()
                                + " PDF(s) encontrado(s)."
                );

                btnEnviar.setDisable(ficheirosPdf.isEmpty());

                if (ficheirosPdf.isEmpty()) {
                    mostrarAviso(
                            "A pasta selecionada não possui ficheiros PDF."
                    );
                }

            } catch (Exception ex) {
                mostrarErro(
                        "Erro ao selecionar a pasta",
                        ex
                );
            }
        });

        /*
        * Efetuar os envios numa thread separada.
        */
        btnEnviar.setOnAction(e -> {

            if (ficheirosPdf.isEmpty()) {
                mostrarAviso(
                        "Selecione primeiro uma pasta com ficheiros PDF."
                );
                return;
            }

            btnEnviar.setDisable(true);
            btnEscolher.setDisable(true);
            txtEstado.clear();

            Thread threadEnvio = new Thread(() -> {

                ListaDupla<RelatorioEnvio> resultados =
                        new ListaDupla<>();

                try {
                    ListaDupla<Cliente> listaClientes =
                            clienteRepository.listar();

                    ListaDupla<String> emailsCc =
                            ccRepository.listar();

                    String corpoEmail =
                            configuracaoRepository.obterCorpoEmail();

                    if (listaClientes.estaVazia()) {
                        Platform.runLater(() ->
                                mostrarAviso(
                                        "Não existem clientes cadastrados."
                                )
                        );
                        return;
                    }

                    for (Cliente cliente : listaClientes) {

                        File pdf = ficheirosPdf.get(
                                cliente.arquivoAnexo()
                        );

                        RelatorioEnvio resultado;

                        /*
                        * Verificar se o ficheiro do cliente existe.
                        */
                        if (pdf == null || !pdf.isFile()) {

                            resultado = new RelatorioEnvio(
                                    cliente.nome(),
                                    cliente.email(),
                                    cliente.cil(),
                                    "IGNORADO",
                                    "Anexo não encontrado: "
                                            + cliente.arquivoAnexo(),
                                    LocalDateTime.now()
                            );

                        } else {

                            try {
                                emailService.enviar(
                                        cliente,
                                        pdf,
                                        emailsCc,
                                        corpoEmail
                                );

                                resultado = new RelatorioEnvio(
                                        cliente.nome(),
                                        cliente.email(),
                                        cliente.cil(),
                                        "SUCESSO",
                                        "E-mail enviado com sucesso.",
                                        LocalDateTime.now()
                                );

                            } catch (Exception erroEnvio) {

                                String mensagemErro =
                                        erroEnvio.getMessage() == null
                                                ? erroEnvio.toString()
                                                : erroEnvio.getMessage();

                                resultado = new RelatorioEnvio(
                                        cliente.nome(),
                                        cliente.email(),
                                        cliente.cil(),
                                        "ERRO",
                                        mensagemErro,
                                        LocalDateTime.now()
                                );
                            }
                        }

                        resultados.add(resultado);

                        /*
                        * Guardar o resultado na base de dados.
                        */
                        try {
                            relatorioRepository.inserir(resultado);
                        } catch (Exception erroRelatorio) {
                            System.err.println(
                                    "Erro ao guardar relatório: "
                                            + erroRelatorio.getMessage()
                            );
                        }

                        RelatorioEnvio resultadoFinal = resultado;

                        /*
                        * Atualizar a interface JavaFX.
                        */
                        Platform.runLater(() -> {

                            txtEstado.appendText(
                                    "["
                                            + resultadoFinal.status()
                                            + "] "
                                            + resultadoFinal.nome()
                                            + " — "
                                            + resultadoFinal.mensagem()
                                            + System.lineSeparator()
                            );

                        });
                    }

                    Platform.runLater(() -> {

                        mostrarInfo(
                                "Processamento concluído.\n"
                                        + "Total processado: "
                                        + resultados.tamanho()
                        );

                    });

                } catch (Exception ex) {

                    Platform.runLater(() ->
                            mostrarErro(
                                    "Erro durante o envio em lote",
                                    ex
                            )
                    );

                } finally {

                    Platform.runLater(() -> {
                        btnEnviar.setDisable(false);
                        btnEscolher.setDisable(false);
                    });
                }

            });

            threadEnvio.setDaemon(true);
            threadEnvio.start();
        });

        HBox linhaPasta = new HBox(
                10,
                btnEscolher,
                lblPasta
        );

        box.getChildren().addAll(
                linhaPasta,
                btnEnviar,
                txtEstado
        );

        VBox.setVgrow(
                txtEstado,
                Priority.ALWAYS
        );

        return box;
    }

    private Node criarRelatorio() {
        VBox box =
            criarPagina("Relatório de Envios");

        DatePicker dataInicio =
            new DatePicker();

        DatePicker dataFim =
            new DatePicker();

        TableView<RelatorioEnvio> tabela =
            new TableView<>();

        String[] titulos = {
            "Data",
            "Nome",
            "Email",
            "CIL",
            "Status",
            "Mensagem"
        };

        for (String titulo : titulos) {
            TableColumn<RelatorioEnvio, String>
                coluna =
                    new TableColumn<>(titulo);

            coluna.setCellValueFactory(dados -> {
                RelatorioEnvio r =
                    dados.getValue();

                String valor =
                    switch (titulo) {
                        case "Data" ->
                            r.dataEnvio().format(
                                DateTimeFormatter
                                    .ofPattern(
                                        "dd/MM/yyyy HH:mm"
                                    )
                            );

                        case "Nome" -> r.nome();
                        case "Email" -> r.email();
                        case "CIL" -> r.cil();
                        case "Status" -> r.status();
                        default -> r.mensagem();
                    };

                return new SimpleStringProperty(valor);
            });

            tabela.getColumns().add(coluna);
        }

        tabela.setColumnResizePolicy(
            TableView
                .CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS
        );

        Runnable carregar = () -> {
            try {
                tabela.setItems(
                    FXCollections.observableArrayList(
                        relatorioRepository.listar(
                            dataInicio.getValue(),
                            dataFim.getValue()
                        )
                    )
                );

            } catch (Exception e) {
                mostrarErro(
                    "Erro ao carregar relatório",
                    e
                );
            }
        };

        Button btnFiltrar =
            new Button("Filtrar");

        btnFiltrar.setOnAction(
            e -> carregar.run()
        );

        Button btnExcel =
            new Button("Exportar Excel");

        btnExcel.setOnAction(e -> {
            FileChooser chooser =
                new FileChooser();

            chooser
                .getExtensionFilters()
                .add(
                    new FileChooser.ExtensionFilter(
                        "Excel",
                        "*.xlsx"
                    )
                );

            chooser.setInitialFileName(
                "relatorio.xlsx"
            );

            File destino =
                chooser.showSaveDialog(
                    root.getScene().getWindow()
                );

            if (destino == null) {
                return;
            }

            try {
                exportService.exportarExcel(
                    new ListaDupla<>(tabela.getItems()),
                    destino
                );

                mostrarInfo(
                    "Excel exportado com sucesso."
                );

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao exportar Excel",
                    ex
                );
            }
        });

        Button btnPdf =
            new Button("Exportar PDF");

        btnPdf.setOnAction(e -> {
            FileChooser chooser =
                new FileChooser();

            chooser
                .getExtensionFilters()
                .add(
                    new FileChooser.ExtensionFilter(
                        "PDF",
                        "*.pdf"
                    )
                );

            chooser.setInitialFileName(
                "relatorio.pdf"
            );

            File destino =
                chooser.showSaveDialog(
                    root.getScene().getWindow()
                );

            if (destino == null) {
                return;
            }

            try {
                exportService.exportarPdf(
                    new ListaDupla<>(tabela.getItems()),
                    destino
                );

                mostrarInfo(
                    "PDF exportado com sucesso."
                );

            } catch (Exception ex) {
                mostrarErro(
                    "Erro ao exportar PDF",
                    ex
                );
            }
        });

        box.getChildren().addAll(
            new HBox(
                10,
                new Label("Início:"),
                dataInicio,
                new Label("Fim:"),
                dataFim,
                btnFiltrar
            ),
            tabela,
            new HBox(
                10,
                btnExcel,
                btnPdf
            )
        );

        VBox.setVgrow(
            tabela,
            Priority.ALWAYS
        );

        carregar.run();
        return box;
    }

    private VBox criarPagina(String titulo) {
        VBox box = new VBox(12);
        box.setPadding(new Insets(20));

        Label label = new Label(titulo);

        label.setStyle(
            "-fx-font-size: 22px;" +
            "-fx-font-weight: bold;"
        );

        box.getChildren().add(label);

        return box;
    }

    private boolean emailValido(
            String email
    ) {
        return Pattern.matches(
            "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+$",
            email
        );
    }

    private void mostrarInfo(String mensagem) {
        new Alert(
            Alert.AlertType.INFORMATION,
            mensagem
        ).showAndWait();
    }

    private void mostrarAviso(String mensagem) {
        new Alert(
            Alert.AlertType.WARNING,
            mensagem
        ).showAndWait();
    }

    private void mostrarErro(
            String titulo,
            Throwable erro
    ) {
        Alert alerta =
            new Alert(Alert.AlertType.ERROR);

        alerta.setTitle(titulo);
        alerta.setHeaderText(titulo);

        alerta.setContentText(
            erro.getMessage() == null
                ? erro.toString()
                : erro.getMessage()
        );

        alerta.showAndWait();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
