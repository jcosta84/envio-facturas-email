package cv.edu.us.envio;

import com.lowagie.text.Document;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.time.format.DateTimeFormatter;
import java.util.List;

public final class ExportService {

    private static final DateTimeFormatter FORMATO =
        DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm");

    public void exportarExcel(
            List<RelatorioEnvio> dados,
            File destino
    ) throws Exception {

        try (Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet =
                workbook.createSheet("Relatório");

            String[] cabecalhos = {
                "Data",
                "Nome",
                "Email",
                "CIL",
                "Status",
                "Mensagem"
            };

            Row cabecalho = sheet.createRow(0);

            for (int i = 0; i < cabecalhos.length; i++) {
                cabecalho
                    .createCell(i)
                    .setCellValue(cabecalhos[i]);
            }

            int numeroLinha = 1;

            for (RelatorioEnvio r : dados) {
                Row linha =
                    sheet.createRow(numeroLinha++);

                linha.createCell(0)
                     .setCellValue(
                         r.dataEnvio().format(FORMATO)
                     );

                linha.createCell(1)
                     .setCellValue(r.nome());

                linha.createCell(2)
                     .setCellValue(r.email());

                linha.createCell(3)
                     .setCellValue(r.cil());

                linha.createCell(4)
                     .setCellValue(r.status());

                linha.createCell(5)
                     .setCellValue(r.mensagem());
            }

            for (int i = 0; i < cabecalhos.length; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream out =
                     new FileOutputStream(destino)) {

                workbook.write(out);
            }
        }
    }

    public void exportarPdf(
            List<RelatorioEnvio> dados,
            File destino
    ) throws Exception {

        Document documento = new Document();

        PdfWriter.getInstance(
            documento,
            new FileOutputStream(destino)
        );

        documento.open();

        documento.add(
            new Paragraph("Relatório de Envios")
        );

        documento.add(new Paragraph(" "));

        PdfPTable tabela = new PdfPTable(5);
        tabela.setWidthPercentage(100);

        String[] cabecalhos = {
            "Data",
            "Nome",
            "Email",
            "CIL",
            "Status"
        };

        for (String cabecalho : cabecalhos) {
            tabela.addCell(cabecalho);
        }

        for (RelatorioEnvio r : dados) {
            tabela.addCell(
                r.dataEnvio().format(FORMATO)
            );

            tabela.addCell(r.nome());
            tabela.addCell(r.email());
            tabela.addCell(r.cil());
            tabela.addCell(r.status());
        }

        documento.add(tabela);
        documento.close();
    }
}
