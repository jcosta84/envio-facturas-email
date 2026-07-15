package cv.edu.us.envio;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Leitor simples de ficheiros INI.
 *
 * O ficheiro config.ini deve ficar na raiz do projeto,
 * ao lado do pom.xml.
 */
public final class ConfigIni {

    private final Map<String, Map<String, String>> secoes = new HashMap<>();

    public ConfigIni(String caminho) {
        carregar(Path.of(caminho));
    }

    private void carregar(Path caminho) {
        if (!Files.exists(caminho)) {
            throw new IllegalStateException(
                "Ficheiro de configuração não encontrado: " + caminho.toAbsolutePath()
            );
        }

        try {
            List<String> linhas = Files.readAllLines(caminho, StandardCharsets.UTF_8);
            String secaoAtual = "";

            for (String linhaOriginal : linhas) {
                String linha = linhaOriginal.trim();

                if (linha.isEmpty() || linha.startsWith("#") || linha.startsWith(";")) {
                    continue;
                }

                if (linha.startsWith("[") && linha.endsWith("]")) {
                    secaoAtual = linha.substring(1, linha.length() - 1)
                                      .trim()
                                      .toUpperCase();
                    secoes.putIfAbsent(secaoAtual, new HashMap<>());
                    continue;
                }

                int separador = linha.indexOf('=');
                if (separador <= 0) {
                    continue;
                }

                String chave = linha.substring(0, separador).trim().toUpperCase();
                String valor = linha.substring(separador + 1).trim();

                secoes.computeIfAbsent(secaoAtual, k -> new HashMap<>())
                      .put(chave, valor);
            }

            validar();
        } catch (IOException e) {
            throw new IllegalStateException("Erro ao ler config.ini.", e);
        }
    }

    private void validar() {
        String[][] obrigatorios = {
            {"DATABASE", "HOST"},
            {"DATABASE", "PORT"},
            {"DATABASE", "DATABASE"},
            {"DATABASE", "USERNAME"},
            {"DATABASE", "PASSWORD"},
            {"EMAIL", "REMETENTE"},
            {"EMAIL", "SENHA_APP"},
            {"EMAIL", "ASSUNTO"}
        };

        for (String[] item : obrigatorios) {
            String valor = get(item[0], item[1]);
            if (valor.isBlank()) {
                throw new IllegalStateException(
                    "Configuração obrigatória vazia: [" +
                    item[0] + "] " + item[1]
                );
            }
        }
    }

    public String get(String secao, String chave) {
        Map<String, String> dados =
            secoes.get(secao.trim().toUpperCase());

        if (dados == null) {
            return "";
        }

        return dados.getOrDefault(
            chave.trim().toUpperCase(),
            ""
        ).trim();
    }

    public int getInt(String secao, String chave) {
        return Integer.parseInt(get(secao, chave));
    }
}
