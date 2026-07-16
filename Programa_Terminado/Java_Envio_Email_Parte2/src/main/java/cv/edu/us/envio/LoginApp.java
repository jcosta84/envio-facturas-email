package cv.edu.us.envio;

import javafx.application.Application;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.control.TextField;
import javafx.scene.input.KeyCode;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

/**
 * Tela de autenticação apresentada antes do sistema principal.
 */
public final class LoginApp extends Application {

    private Stage stage;
    private UsuarioRepository usuarioRepository;

    private TextField txtUsername;
    private PasswordField txtPassword;
    private Label lblMensagem;
    private Button btnEntrar;
    private ProgressIndicator progresso;

    @Override
    public void start(Stage stage) {
        this.stage = stage;

        try {
            ConfigIni config = new ConfigIni("config.ini");
            Database database = new Database(config);
            database.testar();

            usuarioRepository =
                new UsuarioRepository(database);

            mostrarTelaLogin();

        } catch (Exception e) {
            mostrarErro(
                "Erro ao iniciar",
                e.getMessage() == null
                    ? e.toString()
                    : e.getMessage()
            );

            Platform.exit();
        }
    }

    private void mostrarTelaLogin() {
        BorderPane root = new BorderPane();
        root.getStyleClass().add("login-root");

        VBox painelEsquerdo = criarPainelEsquerdo();
        VBox painelLogin = criarPainelLogin();

        StackPane areaLogin = new StackPane(painelLogin);
        areaLogin.getStyleClass().add("login-area");

        root.setLeft(painelEsquerdo);
        root.setCenter(areaLogin);

        Scene scene = new Scene(root, 1000, 650);

        var css = getClass().getResource("/login.css");
        if (css != null) {
            scene.getStylesheets().add(
                css.toExternalForm()
            );
        }

        scene.setOnKeyPressed(event -> {
            if (event.getCode() == KeyCode.ENTER) {
                autenticar();
            }
        });

        stage.setTitle(
            "Login — Sistema de Envio de E-mails"
        );

        stage.setMinWidth(900);
        stage.setMinHeight(600);
        stage.setScene(scene);
        stage.centerOnScreen();
        stage.show();

        Platform.runLater(
            () -> txtUsername.requestFocus()
        );
    }

    private VBox criarPainelEsquerdo() {
        Label icone = new Label("✉");
        icone.getStyleClass().add("login-brand-icon");

        Label titulo =
            new Label("Envio de E-mails");

        titulo.getStyleClass().add(
            "login-brand-title"
        );

        Label subtitulo = new Label(
            "Gestão de clientes, documentos PDF,\n" +
            "envios e relatórios numa única aplicação."
        );

        subtitulo.getStyleClass().add(
            "login-brand-subtitle"
        );

        VBox recursos = new VBox(
            16,
            criarRecurso(
                "✓",
                "Gestão de clientes"
            ),
            criarRecurso(
                "✓",
                "Envio automático de PDFs"
            ),
            criarRecurso(
                "✓",
                "Relatórios em Excel e PDF"
            )
        );

        recursos.getStyleClass().add(
            "login-resources"
        );

        Region separador = new Region();
        VBox.setVgrow(
            separador,
            Priority.ALWAYS
        );

        Label versao =
            new Label("Sistema v1.0");

        versao.getStyleClass().add(
            "login-version"
        );

        VBox painel = new VBox(
            18,
            icone,
            titulo,
            subtitulo,
            recursos,
            separador,
            versao
        );

        painel.setPadding(
            new Insets(60, 45, 35, 45)
        );

        painel.setPrefWidth(430);
        painel.getStyleClass().add(
            "login-side-panel"
        );

        return painel;
    }

    private HBox criarRecurso(
            String simbolo,
            String texto
    ) {
        Label lblSimbolo =
            new Label(simbolo);

        lblSimbolo.getStyleClass().add(
            "login-feature-icon"
        );

        Label lblTexto =
            new Label(texto);

        lblTexto.getStyleClass().add(
            "login-feature-text"
        );

        HBox linha =
            new HBox(
                12,
                lblSimbolo,
                lblTexto
            );

        linha.setAlignment(
            Pos.CENTER_LEFT
        );

        return linha;
    }

    private VBox criarPainelLogin() {
        Label titulo =
            new Label("Bem-vindo");

        titulo.getStyleClass().add(
            "login-title"
        );

        Label descricao = new Label(
            "Introduza os seus dados para continuar."
        );

        descricao.getStyleClass().add(
            "login-description"
        );

        Label lblUsername =
            new Label("Utilizador");

        lblUsername.getStyleClass().add(
            "login-field-label"
        );

        txtUsername =
            new TextField();

        txtUsername.setPromptText(
            "Nome de utilizador"
        );

        txtUsername.getStyleClass().add(
            "login-field"
        );

        Label lblPassword =
            new Label("Palavra-passe");

        lblPassword.getStyleClass().add(
            "login-field-label"
        );

        txtPassword =
            new PasswordField();

        txtPassword.setPromptText(
            "Palavra-passe"
        );

        txtPassword.getStyleClass().add(
            "login-field"
        );

        lblMensagem = new Label();
        lblMensagem.setWrapText(true);
        lblMensagem.getStyleClass().add(
            "login-message"
        );

        progresso =
            new ProgressIndicator();

        progresso.setMaxSize(22, 22);
        progresso.setVisible(false);

        btnEntrar =
            new Button("Entrar");

        btnEntrar.setMaxWidth(
            Double.MAX_VALUE
        );

        btnEntrar.getStyleClass().add(
            "login-button"
        );

        btnEntrar.setOnAction(
            event -> autenticar()
        );

        HBox linhaBotao =
            new HBox(
                12,
                btnEntrar,
                progresso
            );

        linhaBotao.setAlignment(
            Pos.CENTER
        );

        HBox.setHgrow(
            btnEntrar,
            Priority.ALWAYS
        );

        VBox painel =
            new VBox(
                12,
                titulo,
                descricao,
                criarEspaco(10),
                lblUsername,
                txtUsername,
                lblPassword,
                txtPassword,
                lblMensagem,
                linhaBotao
            );

        painel.setMaxWidth(390);
        painel.setPadding(
            new Insets(40)
        );

        painel.getStyleClass().add(
            "login-card"
        );

        return painel;
    }

    private Region criarEspaco(
            double altura
    ) {
        Region espaco = new Region();
        espaco.setMinHeight(altura);
        return espaco;
    }

    private void autenticar() {
        String username =
            txtUsername.getText().trim();

        String password =
            txtPassword.getText();

        lblMensagem.setText("");

        if (username.isBlank()
                || password.isBlank()) {

            lblMensagem.setText(
                "Preencha o utilizador " +
                "e a palavra-passe."
            );

            return;
        }

        alterarEstadoLogin(true);

        Thread thread =
            new Thread(() -> {
                try {
                    Usuario usuario =
                        usuarioRepository.autenticar(
                            username,
                            password
                        );

                    Platform.runLater(() -> {
                        alterarEstadoLogin(false);

                        if (usuario == null) {
                            lblMensagem.setText(
                                "Utilizador ou " +
                                "palavra-passe incorretos."
                            );

                            txtPassword.clear();
                            txtPassword.requestFocus();
                            return;
                        }

                        abrirSistema(usuario);
                    });

                } catch (Exception e) {
                    Platform.runLater(() -> {
                        alterarEstadoLogin(false);

                        mostrarErro(
                            "Erro de autenticação",
                            e.getMessage() == null
                                ? e.toString()
                                : e.getMessage()
                        );
                    });
                }
            });

        thread.setDaemon(true);
        thread.start();
    }

    private void alterarEstadoLogin(
            boolean carregando
    ) {
        btnEntrar.setDisable(carregando);
        txtUsername.setDisable(carregando);
        txtPassword.setDisable(carregando);
        progresso.setVisible(carregando);

        btnEntrar.setText(
            carregando
                ? "A autenticar..."
                : "Entrar"
        );
    }

    private void abrirSistema(
            Usuario usuario
    ) {
        try {
            stage.setUserData(usuario);

            MainApp mainApp =
                new MainApp();

            mainApp.start(stage);

        } catch (Exception e) {
            mostrarErro(
                "Erro ao abrir o sistema",
                e.getMessage() == null
                    ? e.toString()
                    : e.getMessage()
            );
        }
    }

    private void mostrarErro(
            String titulo,
            String mensagem
    ) {
        Alert alerta =
            new Alert(
                Alert.AlertType.ERROR
            );

        alerta.setTitle(titulo);
        alerta.setHeaderText(titulo);
        alerta.setContentText(mensagem);
        alerta.showAndWait();
    }

    public static void main(
            String[] args
    ) {
        launch(args);
    }
}
