package cv.edu.us.envio;

import java.time.LocalDateTime;

public record RelatorioEnvio(
        String nome,
        String email,
        String cil,
        String status,
        String mensagem,
        LocalDateTime dataEnvio
) {}
