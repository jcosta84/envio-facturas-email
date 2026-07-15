package cv.edu.us.envio;

public record Cliente(String cil, String nome, String email, String arquivoAnexo) {
    public Cliente {
        cil = cil == null ? "" : cil.trim();
        nome = nome == null ? "" : nome.trim();
        email = email == null ? "" : email.trim();
        arquivoAnexo = arquivoAnexo == null ? "" : arquivoAnexo.trim();
    }
}
