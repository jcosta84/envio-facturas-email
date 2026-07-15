package cv.edu.us.envio;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.geometry.Insets;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class MainApp extends Application {
    private final BorderPane root = new BorderPane();

    private AppConfig config;
    private Database db;
    private ClienteRepository clientes;
    private CcRepository ccRepo;
    private ConfiguracaoRepository configRepo;
    private RelatorioRepository relatorioRepo;
    private EmailService emailService;
    private final ExportService exportService = new ExportService();

    @Override
    public void start(Stage stage) {
        try {
            config = new AppConfig();
            db = new Database(config);
            db.testar();

            clientes = new ClienteRepository(db);
            ccRepo = new CcRepository(db);
            configRepo = new ConfiguracaoRepository(db);
            relatorioRepo = new RelatorioRepository(db);
            emailService = new EmailService(config);

            root.setLeft(criarMenu());
            root.setCenter(criarCadastro());

            Scene scene = new Scene(root, 1100, 720);
            stage.setTitle("Gestão de Clientes e Envio de E-mails");
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
        menu.setPrefWidth(210);

        Label titulo = new Label("MENU");
        titulo.setStyle("-fx-font-size: 18px; -fx-font-weight: bold;");

        Button cadastro = botaoMenu("Cadastro", () -> root.setCenter(criarCadastro()));
        Button cc = botaoMenu("CC Conhecimento", () -> root.setCenter(criarCc()));
        Button consulta = botaoMenu("Consulta", () -> root.setCenter(criarConsulta()));
        Button corpo = botaoMenu("Corpo do E-mail", () -> root.setCenter(criarCorpoEmail()));
        Button envio = botaoMenu("Enviar E-mails", () -> root.setCenter(criarEnvio()));
        Button relatorio = botaoMenu("Relatório", () -> root.setCenter(criarRelatorio()));

        menu.getChildren().addAll(titulo, cadastro, cc, consulta, corpo, envio, relatorio);
        return menu;
    }

    private Button botaoMenu(String texto, Runnable acao) {
        Button b = new Button(texto);
        b.setMaxWidth(Double.MAX_VALUE);
        b.setOnAction(e -> acao.run());
        return b;
    }

    private Node criarCadastro() {
        VBox box = pagina("Cadastro de Cliente");
        TextField cil = new TextField();
        TextField nome = new TextField();
        TextField email = new TextField();
        Label anexo = new Label();

        cil.textProperty().addListener((obs, antigo, novo) ->
            anexo.setText(novo.isBlank() ? "" : novo.trim() + ".pdf")
        );

        GridPane form = new GridPane();
        form.setHgap(10);
        form.setVgap(10);
        form.addRow(0, new Label("CIL:"), cil);
        form.addRow(1, new Label("Nome:"), nome);
        form.addRow(2, new Label("Email:"), email);
        form.addRow(3, new Label("Nome do PDF:"), anexo);

        Button guardar = new Button("Cadastrar");
        Button limpar = new Button("Limpar");

        guardar.setOnAction(e -> {
            String c = cil.getText().trim();
            String n = nome.getText().trim();
            String em = email.getText().trim();

            if (c.isBlank() || n.isBlank() || em.isBlank()) {
                mostrarAviso("Preencha todos os campos.");
                return;
            }
            if (!emailValido(em)) {
                mostrarAviso("E-mail inválido.");
                return;
            }

            try {
                if (clientes.existeCil(c)) {
                    mostrarAviso("Já existe um cliente com este CIL.");
                    return;
                }
                clientes.inserir(new Cliente(c, n, em, c + ".pdf"));
                mostrarInfo("Cliente cadastrado com sucesso.");
                cil.clear(); nome.clear(); email.clear();
            } catch (Exception ex) {
                mostrarErro("Erro ao cadastrar", ex);
            }
        });

        limpar.setOnAction(e -> {
            cil.clear(); nome.clear(); email.clear();
        });

        HBox botoes = new HBox(10, guardar, limpar);
        box.getChildren().addAll(form, botoes);
        return box;
    }

    private Node criarCc() {
        VBox box = pagina("E-mails para Conhecimento (CC)");
        ListView<String> lista = new ListView<>();

        Runnable carregar = () -> {
            try {
                lista.setItems(FXCollections.observableArrayList(ccRepo.listar()));
            } catch (Exception e) {
                mostrarErro("Erro ao carregar CC", e);
            }
        };

        TextField novo = new TextField();
        novo.setPromptText("novo@email.com");

        Button adicionar = new Button("Adicionar");
        Button eliminar = new Button("Eliminar selecionado");

        adicionar.setOnAction(e -> {
            String email = novo.getText().trim();
            if (!emailValido(email)) {
                mostrarAviso("E-mail inválido.");
                return;
            }
            try {
                ccRepo.inserir(email);
                novo.clear();
                carregar.run();
            } catch (Exception ex) {
                mostrarErro("Erro ao adicionar CC", ex);
            }
        });

        eliminar.setOnAction(e -> {
            String email = lista.getSelectionModel().getSelectedItem();
            if (email == null) {
                mostrarAviso("Selecione um e-mail.");
                return;
            }
            try {
                ccRepo.eliminar(email);
                carregar.run();
            } catch (Exception ex) {
                mostrarErro("Erro ao eliminar CC", ex);
            }
        });

        box.getChildren().addAll(lista, new HBox(10, novo, adicionar, eliminar));
        VBox.setVgrow(lista, Priority.ALWAYS);
        carregar.run();
        return box;
    }

    private Node criarConsulta() {
        VBox box = pagina("Consulta e CRUD de Clientes");
        TextField filtro = new TextField();
        filtro.setPromptText("Filtrar por CIL ou nome");

        TableView<Cliente> tabela = criarTabelaClientes();
        Runnable carregar = () -> {
            try {
                tabela.setItems(FXCollections.observableArrayList(clientes.listar()));
            } catch (Exception e) {
                mostrarErro("Erro ao carregar clientes", e);
            }
        };

        filtro.textProperty().addListener((obs, a, valor) -> {
            try {
                String q = valor.toLowerCase();
                tabela.setItems(FXCollections.observableArrayList(
                    clientes.listar().stream()
                        .filter(c -> c.cil().toLowerCase().contains(q) ||
                                     c.nome().toLowerCase().contains(q))
                        .toList()
                ));
            } catch (Exception e) {
                mostrarErro("Erro ao filtrar", e);
            }
        });

        TextField novoEmail = new TextField();
        novoEmail.setPromptText("Novo e-mail");
        Button atualizar = new Button("Atualizar e-mail");
        Button eliminar = new Button("Eliminar");

        atualizar.setOnAction(e -> {
            Cliente c = tabela.getSelectionModel().getSelectedItem();
            if (c == null || !emailValido(novoEmail.getText().trim())) {
                mostrarAviso("Selecione um cliente e informe um e-mail válido.");
                return;
            }
            try {
                clientes.atualizarEmail(c.cil(), novoEmail.getText());
                carregar.run();
            } catch (Exception ex) {
                mostrarErro("Erro ao atualizar", ex);
            }
        });

        eliminar.setOnAction(e -> {
            Cliente c = tabela.getSelectionModel().getSelectedItem();
            if (c == null) {
                mostrarAviso("Selecione um cliente.");
                return;
            }
            try {
                clientes.eliminar(c.cil());
                carregar.run();
            } catch (Exception ex) {
                mostrarErro("Erro ao eliminar", ex);
            }
        });

        box.getChildren().addAll(filtro, tabela, new HBox(10, novoEmail, atualizar, eliminar));
        VBox.setVgrow(tabela, Priority.ALWAYS);
        carregar.run();
        return box;
    }

    private TableView<Cliente> criarTabelaClientes() {
        TableView<Cliente> tabela = new TableView<>();
        TableColumn<Cliente, String> cil = new TableColumn<>("CIL");
        cil.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().cil()));
        TableColumn<Cliente, String> nome = new TableColumn<>("Nome");
        nome.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().nome()));
        TableColumn<Cliente, String> email = new TableColumn<>("Email");
        email.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().email()));
        TableColumn<Cliente, String> anexo = new TableColumn<>("Anexo");
        anexo.setCellValueFactory(d -> new SimpleStringProperty(d.getValue().arquivoAnexo()));
        tabela.getColumns().addAll(cil, nome, email, anexo);
        tabela.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);
        return tabela;
    }

    private Node criarCorpoEmail() {
        VBox box = pagina("Corpo do E-mail");
        TextArea texto = new TextArea();
        Button carregar = new Button("Recarregar");
        Button guardar = new Button("Guardar alterações");

        Runnable recarregar = () -> {
            try {
                texto.setText(configRepo.obterCorpoEmail());
            } catch (Exception e) {
                mostrarErro("Erro ao carregar o corpo", e);
            }
        };

        carregar.setOnAction(e -> recarregar.run());
        guardar.setOnAction(e -> {
            if (texto.getText().isBlank()) {
                mostrarAviso("O corpo do e-mail não pode ficar vazio.");
                return;
            }
            try {
                configRepo.guardarCorpoEmail(texto.getText().trim());
                mostrarInfo("Corpo do e-mail atualizado.");
            } catch (Exception ex) {
                mostrarErro("Erro ao guardar", ex);
            }
        });

        box.getChildren().addAll(texto, new HBox(10, guardar, carregar));
        VBox.setVgrow(texto, Priority.ALWAYS);
        recarregar.run();
        return box;
    }

    private Node criarEnvio() {
        VBox box = pagina("Envio de E-mails com PDF");
        Label pastaLabel = new Label("Nenhuma pasta selecionada.");
        TextArea estado = new TextArea();
        estado.setEditable(false);

        Map<String, File> pdfs = new HashMap<>();

        Button escolher = new Button("Selecionar diretório");
        escolher.setOnAction(e -> {
            DirectoryChooser dc = new DirectoryChooser();
            File dir = dc.showDialog(root.getScene().getWindow());
            if (dir == null) return;

            pdfs.clear();
            File[] ficheiros = dir.listFiles((d, nome) -> nome.toLowerCase().endsWith(".pdf"));
            if (ficheiros != null) {
                for (File f : ficheiros) pdfs.put(f.getName(), f);
            }
            pastaLabel.setText(pdfs.size() + " PDF(s) encontrado(s).");
        });

        Button enviar = new Button("Enviar em lote");
        enviar.setOnAction(e -> {
            if (pdfs.isEmpty()) {
                mostrarAviso("Selecione uma pasta com ficheiros PDF.");
                return;
            }
            enviar.setDisable(true);
            estado.clear();

            Thread.ofVirtual().start(() -> {
                ListaLigada<RelatorioEnvio> resultados = new ListaLigada<>();
                try {
                    List<Cliente> listaClientes = clientes.listar();
                    List<String> cc = ccRepo.listar();
                    String corpo = configRepo.obterCorpoEmail();

                    for (Cliente c : listaClientes) {
                        File pdf = pdfs.get(c.arquivoAnexo());
                        RelatorioEnvio r;

                        if (pdf == null || !pdf.isFile()) {
                            r = new RelatorioEnvio(
                                c.nome(), c.email(), c.cil(), "IGNORADO",
                                "Anexo não encontrado: " + c.arquivoAnexo(),
                                LocalDateTime.now()
                            );
                        } else {
                            try {
                                emailService.enviar(c, pdf, cc, corpo);
                                r = new RelatorioEnvio(
                                    c.nome(), c.email(), c.cil(), "SUCESSO",
                                    "E-mail enviado com sucesso.",
                                    LocalDateTime.now()
                                );
                            } catch (Exception ex) {
                                r = new RelatorioEnvio(
                                    c.nome(), c.email(), c.cil(), "ERRO",
                                    ex.getMessage(),
                                    LocalDateTime.now()
                                );
                            }
                        }

                        resultados.adicionar(r);
                        relatorioRepo.inserir(r);
                        RelatorioEnvio finalR = r;
                        Platform.runLater(() ->
                            estado.appendText("[" + finalR.status() + "] " +
                                    finalR.nome() + " - " + finalR.mensagem() + "\n")
                        );
                    }

                    Platform.runLater(() -> mostrarInfo(
                        "Processamento concluído. Total: " + resultados.tamanho()
                    ));
                } catch (Exception ex) {
                    Platform.runLater(() -> mostrarErro("Erro no envio em lote", ex));
                } finally {
                    Platform.runLater(() -> enviar.setDisable(false));
                }
            });
        });

        box.getChildren().addAll(new HBox(10, escolher, pastaLabel), enviar, estado);
        VBox.setVgrow(estado, Priority.ALWAYS);
        return box;
    }

    private Node criarRelatorio() {
        VBox box = pagina("Relatório de Envios");
        DatePicker inicio = new DatePicker();
        DatePicker fim = new DatePicker();
        TableView<RelatorioEnvio> tabela = new TableView<>();

        String[] titulos = {"Data", "Nome", "Email", "CIL", "Status", "Mensagem"};
        for (String titulo : titulos) {
            TableColumn<RelatorioEnvio, String> col = new TableColumn<>(titulo);
            col.setCellValueFactory(d -> {
                RelatorioEnvio r = d.getValue();
                return new SimpleStringProperty(switch (titulo) {
                    case "Data" -> r.dataEnvio().format(DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm"));
                    case "Nome" -> r.nome();
                    case "Email" -> r.email();
                    case "CIL" -> r.cil();
                    case "Status" -> r.status();
                    default -> r.mensagem();
                });
            });
            tabela.getColumns().add(col);
        }
        tabela.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY_ALL_COLUMNS);

        Runnable carregar = () -> {
            try {
                tabela.setItems(FXCollections.observableArrayList(
                    relatorioRepo.listar(inicio.getValue(), fim.getValue())
                ));
            } catch (Exception e) {
                mostrarErro("Erro ao carregar relatório", e);
            }
        };

        Button filtrar = new Button("Filtrar");
        filtrar.setOnAction(e -> carregar.run());

        Button excel = new Button("Exportar Excel");
        excel.setOnAction(e -> {
            FileChooser fc = new FileChooser();
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel", "*.xlsx"));
            fc.setInitialFileName("relatorio.xlsx");
            File f = fc.showSaveDialog(root.getScene().getWindow());
            if (f == null) return;
            try {
                exportService.exportarExcel(new ArrayList<>(tabela.getItems()), f);
                mostrarInfo("Excel exportado.");
            } catch (Exception ex) {
                mostrarErro("Erro ao exportar Excel", ex);
            }
        });

        Button pdf = new Button("Exportar PDF");
        pdf.setOnAction(e -> {
            FileChooser fc = new FileChooser();
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF", "*.pdf"));
            fc.setInitialFileName("relatorio.pdf");
            File f = fc.showSaveDialog(root.getScene().getWindow());
            if (f == null) return;
            try {
                exportService.exportarPdf(new ArrayList<>(tabela.getItems()), f);
                mostrarInfo("PDF exportado.");
            } catch (Exception ex) {
                mostrarErro("Erro ao exportar PDF", ex);
            }
        });

        box.getChildren().addAll(
            new HBox(10, new Label("Início:"), inicio, new Label("Fim:"), fim, filtrar),
            tabela,
            new HBox(10, excel, pdf)
        );
        VBox.setVgrow(tabela, Priority.ALWAYS);
        carregar.run();
        return box;
    }

    private VBox pagina(String titulo) {
        VBox box = new VBox(12);
        box.setPadding(new Insets(20));
        Label label = new Label(titulo);
        label.setStyle("-fx-font-size: 22px; -fx-font-weight: bold;");
        box.getChildren().add(label);
        return box;
    }

    private boolean emailValido(String email) {
        return Pattern.matches("^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+$", email);
    }

    private void mostrarInfo(String msg) {
        new Alert(Alert.AlertType.INFORMATION, msg).showAndWait();
    }

    private void mostrarAviso(String msg) {
        new Alert(Alert.AlertType.WARNING, msg).showAndWait();
    }

    private void mostrarErro(String titulo, Throwable e) {
        Alert a = new Alert(Alert.AlertType.ERROR);
        a.setTitle(titulo);
        a.setHeaderText(titulo);
        a.setContentText(e.getMessage() == null ? e.toString() : e.getMessage());
        a.showAndWait();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
