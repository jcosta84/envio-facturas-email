package cv.edu.us.envio;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.ComboBox;
import javafx.scene.control.DatePicker;
import javafx.scene.control.Label;
import javafx.scene.control.Separator;
import javafx.scene.control.ListView;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.scene.control.PasswordField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.util.Duration;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

public class MainApp extends Application {

    private final BorderPane root = new BorderPane();
    private final Label lblData = new Label();
    private final Label lblHora = new Label();
    private Button botaoAtivo;

    private ConfigIni config;
    private Database database;
    private ClienteRepository clienteRepository;
    private CcRepository ccRepository;
    private ConfiguracaoRepository configuracaoRepository;
    private RelatorioRepository relatorioRepository;
    private EmailService emailService;
    private UsuarioRepository usuarioRepository;
    private Usuario usuarioAtual;
    private Timeline relogio;

    private final ExportService exportService =
        new ExportService();

    @Override
    public void start(Stage stage) {
        try {
            config = new ConfigIni("config.ini");

            database = new Database(config);
            database.testar();

            clienteRepository = new ClienteRepository(database);
            ccRepository = new CcRepository(database);
            configuracaoRepository = new ConfiguracaoRepository(database);
            relatorioRepository = new RelatorioRepository(database);
            emailService = new EmailService(config);
            usuarioRepository = new UsuarioRepository(database);

            Object dadosSessao = stage.getUserData();
            if (!(dadosSessao instanceof Usuario)) {
                new LoginApp().start(stage);
                return;
            }
            usuarioAtual = (Usuario) dadosSessao;

            root.getStyleClass().add("app-root");
            root.setLeft(criarMenu());
            root.setCenter(criarInicio());
            root.setBottom(criarRodape());

            Scene scene = new Scene(root, 1280, 820);
            scene.getStylesheets().add(
                getClass().getResource("/app.css").toExternalForm()
            );

            stage.setTitle("Gestão e Envio de E-mails");
            stage.setMinWidth(1050);
            stage.setMinHeight(700);
            stage.setScene(scene);
            stage.show();

            iniciarRelogio();

        } catch (Exception e) {
            mostrarErro("Erro ao iniciar", e);
            Platform.exit();
        }
    }

    private Node criarMenu() {
        VBox menu = new VBox(12);
        menu.getStyleClass().add("sidebar");
        menu.setPadding(new Insets(28, 18, 22, 18));
        menu.setPrefWidth(245);
        menu.setMinWidth(225);

        Label utilizador = new Label(usuarioAtual.username());
        utilizador.getStyleClass().add("sidebar-user");

        Label nivel = new Label(usuarioAtual.nivel().toUpperCase());
        nivel.getStyleClass().add("sidebar-role");

        VBox identificacao = new VBox(2, utilizador, nivel);
        identificacao.setAlignment(Pos.CENTER);
        identificacao.setPadding(new Insets(0, 0, 18, 0));

        Button inicio = criarBotaoMenu("⌂", "Início", this::criarInicio);
        Button clientes = criarBotaoMenu("👤", "Clientes", this::criarCadastro);
        Button consulta = criarBotaoMenu("⌕", "Consulta", this::criarConsulta);
        Button envio = criarBotaoMenu("✉", "Enviar E-mails", this::criarEnvio);
        Button corpoEmail = criarBotaoMenu("⚙", "Corpo do E-mail", this::criarCorpoEmail);
        Button cc = criarBotaoMenu("👥", "Gerir CC", this::criarCc);
        Button relatorio = criarBotaoMenu("▥", "Relatórios", this::criarRelatorio);
        Button definicoes = criarBotaoMenu("🔒", "Definições", this::criarDefinicoes);

        Button utilizadores = null;
        if (usuarioAtual.isAdmin()) {
            utilizadores = criarBotaoMenu(
                "🛡",
                "Utilizadores",
                this::criarGestaoUtilizadores
            );
        }

        Region espaco = new Region();
        VBox.setVgrow(espaco, Priority.ALWAYS);

        Separator separador = new Separator();
        separador.getStyleClass().add("sidebar-separator");

        Button terminar = criarBotaoMenu(
            "⏻",
            "Terminar Sessão",
            this::terminarSessao
        );

        Button sair = criarBotaoMenu(
            "↪",
            "Sair",
            this::sairAplicacao
        );

        menu.getChildren().addAll(
            identificacao,
            inicio,
            clientes,
            consulta,
            envio,
            corpoEmail,
            cc,
            relatorio
        );

        if (utilizadores != null) {
            menu.getChildren().add(utilizadores);
        }

        menu.getChildren().addAll(
            definicoes,
            espaco,
            separador,
            terminar,
            sair
        );

        Platform.runLater(() -> definirBotaoAtivo(inicio));
        return menu;
    }

    private Button criarBotaoMenu(
            String icone,
            String texto,
            java.util.function.Supplier<Node> pagina
    ) {
        Button botao = new Button(icone + "   " + texto);
        botao.getStyleClass().add("sidebar-button");
        botao.setMaxWidth(Double.MAX_VALUE);
        botao.setAlignment(Pos.CENTER_LEFT);
        botao.setOnAction(e -> navegar(botao, pagina));
        return botao;
    }

    private Button criarBotaoMenu(
            String icone,
            String texto,
            Runnable acao
    ) {
        Button botao = new Button(icone + "   " + texto);
        botao.getStyleClass().add("sidebar-button");
        botao.setMaxWidth(Double.MAX_VALUE);
        botao.setAlignment(Pos.CENTER_LEFT);
        botao.setOnAction(e -> acao.run());
        return botao;
    }

    private void navegar(
            Button origem,
            java.util.function.Supplier<Node> pagina
    ) {
        definirBotaoAtivo(origem);
        root.setCenter(pagina.get());
    }

    private void definirBotaoAtivo(Button botao) {
        if (botaoAtivo != null) {
            botaoAtivo.getStyleClass().remove("sidebar-button-active");
        }
        botaoAtivo = botao;
        if (!botao.getStyleClass().contains("sidebar-button-active")) {
            botao.getStyleClass().add("sidebar-button-active");
        }
    }

    private Node criarInicio() {
        VBox pagina = criarPagina("Início");

        Label subtitulo = new Label(
            "Bem-vindo(a) ao sistema de gestão de clientes e envio de e-mails."
        );
        subtitulo.getStyleClass().add("page-subtitle");

        HBox cards = new HBox(16);
        cards.setFillHeight(true);

        try {
            int clientes = clienteRepository.listar().tamanho();
            int cc = ccRepository.listar().tamanho();
            int relatorios = relatorioRepository.listar(null, null).tamanho();
            long sucesso = relatorioRepository.listar(null, null)
                .stream()
                .filter(r -> "SUCESSO".equalsIgnoreCase(r.status()))
                .count();

            cards.getChildren().addAll(
                criarCardResumo("👤", String.valueOf(clientes), "Clientes", "Cadastrados"),
                criarCardResumo("✉", String.valueOf(sucesso), "E-mails", "Enviados"),
                criarCardResumo("▥", String.valueOf(relatorios), "Relatórios", "Registados"),
                criarCardResumo("👥", String.valueOf(cc), "Destinatários CC", "Ativos")
            );
        } catch (Exception e) {
            cards.getChildren().addAll(
                criarCardResumo("👤", "—", "Clientes", "Indisponível"),
                criarCardResumo("✉", "—", "E-mails", "Indisponível"),
                criarCardResumo("▥", "—", "Relatórios", "Indisponível"),
                criarCardResumo("👥", "—", "Destinatários CC", "Indisponível")
            );
        }

        for (Node card : cards.getChildren()) {
            HBox.setHgrow(card, Priority.ALWAYS);
        }

        VBox atividade = new VBox(12);
        atividade.getStyleClass().add("panel-card");

        Label tituloAtividade = new Label("Atividade recente");
        tituloAtividade.getStyleClass().add("section-title");

        TableView<RelatorioEnvio> tabela = criarTabelaRelatorioResumida();
        tabela.setPrefHeight(260);

        try {
            ListaDupla<RelatorioEnvio> todos = relatorioRepository.listar(null, null);
            ListaDupla<RelatorioEnvio> recentes = new ListaDupla<>();
            for (int i = 0; i < Math.min(8, todos.tamanho()); i++) {
                recentes.add(todos.get(i));
            }
            tabela.setItems(FXCollections.observableArrayList(recentes));
        } catch (Exception e) {
            tabela.setPlaceholder(new Label("Não foi possível carregar os registos."));
        }

        Button verRelatorios = new Button("▥  Ver Relatórios");
        verRelatorios.getStyleClass().add("secondary-button");
        verRelatorios.setOnAction(e -> root.setCenter(criarRelatorio()));

        HBox linhaBotao = new HBox(verRelatorios);
        linhaBotao.setAlignment(Pos.CENTER_RIGHT);

        atividade.getChildren().addAll(tituloAtividade, tabela, linhaBotao);
        VBox.setVgrow(tabela, Priority.ALWAYS);

        VBox acoes = new VBox(12);
        acoes.getStyleClass().add("panel-card");

        Label tituloAcoes = new Label("Ações rápidas");
        tituloAcoes.getStyleClass().add("section-title");

        HBox botoes = new HBox(14,
            criarAcaoRapida("👤", "Novo Cliente", this::criarCadastro),
            criarAcaoRapida("✉", "Enviar E-mails", this::criarEnvio),
            criarAcaoRapida("▥", "Relatório", this::criarRelatorio),
            criarAcaoRapida("⚙", "Corpo do E-mail", this::criarCorpoEmail),
            criarAcaoRapida("👥", "Gerir CC", this::criarCc)
        );

        for (Node botao : botoes.getChildren()) {
            HBox.setHgrow(botao, Priority.ALWAYS);
        }

        acoes.getChildren().addAll(tituloAcoes, botoes);

        pagina.getChildren().addAll(subtitulo, cards, atividade, acoes);
        VBox.setVgrow(atividade, Priority.ALWAYS);
        return pagina;
    }

    private VBox criarCardResumo(
            String icone,
            String numero,
            String titulo,
            String detalhe
    ) {
        Label lblIcone = new Label(icone);
        lblIcone.getStyleClass().add("summary-icon");

        Label lblNumero = new Label(numero);
        lblNumero.getStyleClass().add("summary-number");

        Label lblTitulo = new Label(titulo);
        lblTitulo.getStyleClass().add("summary-title");

        Label lblDetalhe = new Label(detalhe);
        lblDetalhe.getStyleClass().add("summary-detail");

        VBox card = new VBox(8, lblIcone, lblNumero, lblTitulo, lblDetalhe);
        card.getStyleClass().add("summary-card");
        card.setAlignment(Pos.CENTER_LEFT);
        card.setMaxWidth(Double.MAX_VALUE);
        return card;
    }

    private Button criarAcaoRapida(
            String icone,
            String texto,
            java.util.function.Supplier<Node> pagina
    ) {
        Button botao = new Button(icone + "" + texto);
        botao.getStyleClass().add("quick-action");
        botao.setMaxWidth(Double.MAX_VALUE);
        botao.setMaxHeight(Double.MAX_VALUE);
        botao.setOnAction(e -> root.setCenter(pagina.get()));
        return botao;
    }

    private TableView<RelatorioEnvio> criarTabelaRelatorioResumida() {
        TableView<RelatorioEnvio> tabela = new TableView<>();
        tabela.getStyleClass().add("dashboard-table");

        TableColumn<RelatorioEnvio, String> data = new TableColumn<>("Data/Hora");
        data.setCellValueFactory(d -> new SimpleStringProperty(
            d.getValue().dataEnvio().format(
                DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")
            )
        ));

        TableColumn<RelatorioEnvio, String> cliente = new TableColumn<>("Cliente");
        cliente.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().nome()));

        TableColumn<RelatorioEnvio, String> email = new TableColumn<>("E-mail");
        email.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().email()));

        TableColumn<RelatorioEnvio, String> status = new TableColumn<>("Status");
        status.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().status()));

        TableColumn<RelatorioEnvio, String> mensagem = new TableColumn<>("Mensagem");
        mensagem.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().mensagem()));

        tabela.getColumns().addAll(data, cliente, email, status, mensagem);
        tabela.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        tabela.setPlaceholder(new Label("Nenhum registo encontrado."));
        return tabela;
    }

    private Node criarRodape() {
        BorderPane rodape = new BorderPane();
        rodape.getStyleClass().add("status-bar");
        rodape.setPadding(new Insets(10, 18, 10, 18));

        Label estado = new Label("●  Ligação com a base de dados: OK");
        estado.getStyleClass().add("database-status");

        HBox dataHora = new HBox(18,
            new Label("▣"),
            lblData,
            new Label("◷"),
            lblHora
        );
        dataHora.setAlignment(Pos.CENTER_RIGHT);
        dataHora.getStyleClass().add("date-time-box");

        rodape.setLeft(estado);
        rodape.setRight(dataHora);
        atualizarDataHora();
        return rodape;
    }

    private void iniciarRelogio() {
        relogio = new Timeline(
            new KeyFrame(Duration.ZERO, e -> atualizarDataHora()),
            new KeyFrame(Duration.seconds(1))
        );
        relogio.setCycleCount(Timeline.INDEFINITE);
        relogio.play();
    }

    private void atualizarDataHora() {
        LocalDateTime agora = LocalDateTime.now();
        lblData.setText(agora.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        lblHora.setText(agora.format(DateTimeFormatter.ofPattern("HH:mm:ss")));
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


    /**
     * Gestão de utilizadores disponível exclusivamente para administradores.
     */
    private Node criarGestaoUtilizadores() {
        if (!usuarioAtual.isAdmin()) {
            mostrarAviso(
                "Apenas administradores podem gerir utilizadores."
            );
            return criarInicio();
        }

        VBox box = criarPagina("Gestão de Utilizadores");

        Label descricao = new Label(
            "Crie contas de administrador ou gerente. " +
            "Os gerentes não têm acesso a esta página."
        );
        descricao.getStyleClass().add("page-subtitle");

        TextField txtUsername = new TextField();
        txtUsername.setPromptText("Nome de utilizador");

        PasswordField txtPassword = new PasswordField();
        txtPassword.setPromptText("Palavra-passe");

        PasswordField txtConfirmar = new PasswordField();
        txtConfirmar.setPromptText("Confirmar palavra-passe");

        ComboBox<String> cmbNivel = new ComboBox<>();
        cmbNivel.getItems().addAll("admin", "gerente");
        cmbNivel.setValue("gerente");
        cmbNivel.setMaxWidth(Double.MAX_VALUE);

        GridPane formulario = new GridPane();
        formulario.setHgap(12);
        formulario.setVgap(12);
        formulario.addRow(0, new Label("Utilizador:"), txtUsername);
        formulario.addRow(1, new Label("Palavra-passe:"), txtPassword);
        formulario.addRow(2, new Label("Confirmar:"), txtConfirmar);
        formulario.addRow(3, new Label("Nível:"), cmbNivel);

        GridPane.setHgrow(txtUsername, Priority.ALWAYS);
        GridPane.setHgrow(txtPassword, Priority.ALWAYS);
        GridPane.setHgrow(txtConfirmar, Priority.ALWAYS);
        GridPane.setHgrow(cmbNivel, Priority.ALWAYS);

        TableView<Usuario> tabela = criarTabelaUtilizadores();

        Runnable carregar = () -> {
            try {
                tabela.setItems(
                    FXCollections.observableArrayList(
                        usuarioRepository.listar()
                    )
                );
            } catch (Exception e) {
                mostrarErro(
                    "Erro ao carregar utilizadores",
                    e
                );
            }
        };

        Button btnCriar = new Button("Criar utilizador");
        Button btnLimpar = new Button("Limpar");
        Button btnEliminar = new Button("Eliminar selecionado");

        btnCriar.setOnAction(event -> {
            String username = txtUsername.getText().trim();
            String password = txtPassword.getText();
            String confirmar = txtConfirmar.getText();
            String nivel = cmbNivel.getValue();

            if (username.isBlank()
                    || password.isBlank()
                    || confirmar.isBlank()
                    || nivel == null) {

                mostrarAviso("Preencha todos os campos.");
                return;
            }

            if (username.length() < 3) {
                mostrarAviso(
                    "O nome de utilizador deve possuir pelo menos 3 caracteres."
                );
                return;
            }

            if (password.length() < 6) {
                mostrarAviso(
                    "A palavra-passe deve possuir pelo menos 6 caracteres."
                );
                return;
            }

            if (!password.equals(confirmar)) {
                mostrarAviso(
                    "As palavras-passe não coincidem."
                );
                return;
            }

            try {
                usuarioRepository.inserir(
                    username,
                    password,
                    nivel
                );

                mostrarInfo(
                    "Utilizador criado com sucesso."
                );

                txtUsername.clear();
                txtPassword.clear();
                txtConfirmar.clear();
                cmbNivel.setValue("gerente");
                carregar.run();

            } catch (Exception e) {
                mostrarErro(
                    "Erro ao criar utilizador",
                    e
                );
            }
        });

        btnLimpar.setOnAction(event -> {
            txtUsername.clear();
            txtPassword.clear();
            txtConfirmar.clear();
            cmbNivel.setValue("gerente");
        });

        btnEliminar.setOnAction(event -> {
            Usuario selecionado =
                tabela.getSelectionModel().getSelectedItem();

            if (selecionado == null) {
                mostrarAviso(
                    "Selecione um utilizador."
                );
                return;
            }

            if (selecionado.id() == usuarioAtual.id()) {
                mostrarAviso(
                    "Não é possível eliminar o utilizador atualmente autenticado."
                );
                return;
            }

            Alert confirmacao = new Alert(
                Alert.AlertType.CONFIRMATION,
                "Deseja eliminar o utilizador "
                    + selecionado.username() + "?",
                ButtonType.YES,
                ButtonType.NO
            );

            confirmacao.setTitle("Eliminar utilizador");
            confirmacao.setHeaderText(
                "Confirmação de eliminação"
            );

            if (confirmacao.showAndWait().orElse(ButtonType.NO)
                    != ButtonType.YES) {
                return;
            }

            try {
                usuarioRepository.eliminar(
                    selecionado.id()
                );
                carregar.run();
                mostrarInfo(
                    "Utilizador eliminado com sucesso."
                );

            } catch (Exception e) {
                mostrarErro(
                    "Erro ao eliminar utilizador",
                    e
                );
            }
        });

        HBox botoes = new HBox(
            10,
            btnCriar,
            btnLimpar,
            btnEliminar
        );

        box.getChildren().addAll(
            descricao,
            formulario,
            botoes,
            tabela
        );

        VBox.setVgrow(tabela, Priority.ALWAYS);
        carregar.run();

        return box;
    }

    private TableView<Usuario> criarTabelaUtilizadores() {
        TableView<Usuario> tabela = new TableView<>();

        TableColumn<Usuario, String> colId =
            new TableColumn<>("ID");
        colId.setCellValueFactory(
            dados -> new SimpleStringProperty(
                String.valueOf(dados.getValue().id())
            )
        );

        TableColumn<Usuario, String> colUsername =
            new TableColumn<>("Utilizador");
        colUsername.setCellValueFactory(
            dados -> new SimpleStringProperty(
                dados.getValue().username()
            )
        );

        TableColumn<Usuario, String> colNivel =
            new TableColumn<>("Nível");
        colNivel.setCellValueFactory(
            dados -> new SimpleStringProperty(
                dados.getValue().nivel()
            )
        );

        tabela.getColumns().addAll(
            colId,
            colUsername,
            colNivel
        );

        tabela.setColumnResizePolicy(
            TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS
        );

        return tabela;
    }

    /**
     * Página disponível para todos os utilizadores.
     * Cada pessoa altera apenas a sua própria palavra-passe.
     */
    private Node criarDefinicoes() {
        VBox box = criarPagina("Definições");

        Label descricao = new Label(
            "Utilizador autenticado: "
                + usuarioAtual.username()
                + " — "
                + usuarioAtual.nivel()
        );
        descricao.getStyleClass().add("page-subtitle");

        PasswordField txtAtual = new PasswordField();
        txtAtual.setPromptText("Palavra-passe atual");

        PasswordField txtNova = new PasswordField();
        txtNova.setPromptText("Nova palavra-passe");

        PasswordField txtConfirmar = new PasswordField();
        txtConfirmar.setPromptText("Confirmar nova palavra-passe");

        GridPane formulario = new GridPane();
        formulario.setHgap(12);
        formulario.setVgap(12);
        formulario.addRow(
            0,
            new Label("Palavra-passe atual:"),
            txtAtual
        );
        formulario.addRow(
            1,
            new Label("Nova palavra-passe:"),
            txtNova
        );
        formulario.addRow(
            2,
            new Label("Confirmar nova palavra-passe:"),
            txtConfirmar
        );

        GridPane.setHgrow(txtAtual, Priority.ALWAYS);
        GridPane.setHgrow(txtNova, Priority.ALWAYS);
        GridPane.setHgrow(txtConfirmar, Priority.ALWAYS);

        Button btnAlterar =
            new Button("Alterar palavra-passe");

        btnAlterar.setOnAction(event -> {
            String atual = txtAtual.getText();
            String nova = txtNova.getText();
            String confirmar = txtConfirmar.getText();

            if (atual.isBlank()
                    || nova.isBlank()
                    || confirmar.isBlank()) {

                mostrarAviso("Preencha todos os campos.");
                return;
            }

            if (nova.length() < 6) {
                mostrarAviso(
                    "A nova palavra-passe deve possuir pelo menos 6 caracteres."
                );
                return;
            }

            if (!nova.equals(confirmar)) {
                mostrarAviso(
                    "A nova palavra-passe e a confirmação não coincidem."
                );
                return;
            }

            if (atual.equals(nova)) {
                mostrarAviso(
                    "A nova palavra-passe deve ser diferente da atual."
                );
                return;
            }

            try {
                boolean alterada =
                    usuarioRepository.alterarSenhaPropria(
                        usuarioAtual.id(),
                        atual,
                        nova
                    );

                if (!alterada) {
                    mostrarAviso(
                        "A palavra-passe atual está incorreta."
                    );
                    txtAtual.clear();
                    txtAtual.requestFocus();
                    return;
                }

                txtAtual.clear();
                txtNova.clear();
                txtConfirmar.clear();

                mostrarInfo(
                    "Palavra-passe alterada com sucesso."
                );

            } catch (Exception e) {
                mostrarErro(
                    "Erro ao alterar palavra-passe",
                    e
                );
            }
        });

        VBox painel = new VBox(
            16,
            descricao,
            formulario,
            btnAlterar
        );
        painel.getStyleClass().add("panel-card");
        painel.setMaxWidth(650);

        box.getChildren().add(painel);
        return box;
    }

    /**
     * Limpa a sessão atual e reutiliza o mesmo Stage para voltar ao LoginApp.
     */
    private void terminarSessao() {
        Alert confirmacao = new Alert(
            Alert.AlertType.CONFIRMATION,
            "Deseja terminar a sessão atual?",
            ButtonType.YES,
            ButtonType.NO
        );

        confirmacao.setTitle("Terminar sessão");
        confirmacao.setHeaderText(
            "Será redirecionado para a tela de login."
        );

        if (confirmacao.showAndWait().orElse(ButtonType.NO)
                != ButtonType.YES) {
            return;
        }

        try {
            Stage stage =
                (Stage) root.getScene().getWindow();

            if (relogio != null) {
                relogio.stop();
            }

            usuarioAtual = null;
            stage.setUserData(null);

            root.setLeft(null);
            root.setCenter(null);
            root.setBottom(null);

            new LoginApp().start(stage);

        } catch (Exception e) {
            mostrarErro(
                "Erro ao terminar sessão",
                e
            );
        }
    }

    private void sairAplicacao() {
        Alert confirmacao = new Alert(
            Alert.AlertType.CONFIRMATION,
            "Deseja fechar completamente o sistema?",
            ButtonType.YES,
            ButtonType.NO
        );

        confirmacao.setTitle("Sair");
        confirmacao.setHeaderText(
            "Confirmação"
        );

        if (confirmacao.showAndWait().orElse(ButtonType.NO)
                == ButtonType.YES) {

            if (relogio != null) {
                relogio.stop();
            }

            Platform.exit();
        }
    }

    private VBox criarPagina(String titulo) {
        VBox box = new VBox(16);
        box.getStyleClass().add("content-page");
        box.setPadding(new Insets(28));

        Label label = new Label(titulo);
        label.getStyleClass().add("page-title");

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
