package cv.edu.us.envio;

import com.lowagie.text.Document;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.time.format.DateTimeFormatter;
import java.util.List;

public final class ExportService {
    private static final DateTimeFormatter FMT = DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm");

    public void exportarExcel(List<RelatorioEnvio> dados, File destino) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Relatório");
            String[] cab = {"Data", "Nome", "Email", "CIL", "Status", "Mensagem"};
            Row h = sheet.createRow(0);
            for (int i = 0; i < cab.length; i++) h.createCell(i).setCellValue(cab[i]);

            int linha = 1;
            for (RelatorioEnvio r : dados) {
                Row row = sheet.createRow(linha++);
                row.createCell(0).setCellValue(r.dataEnvio().format(FMT));
                row.createCell(1).setCellValue(r.nome());
                row.createCell(2).setCellValue(r.email());
                row.createCell(3).setCellValue(r.cil());
                row.createCell(4).setCellValue(r.status());
                row.createCell(5).setCellValue(r.mensagem());
            }
            for (int i = 0; i < cab.length; i++) sheet.autoSizeColumn(i);

            try (FileOutputStream out = new FileOutputStream(destino)) {
                wb.write(out);
            }
        }
    }

    public void exportarPdf(List<RelatorioEnvio> dados, File destino) throws Exception {
        Document doc = new Document();
        PdfWriter.getInstance(doc, new FileOutputStream(destino));
        doc.open();
        doc.add(new Paragraph("Relatório de Envios"));
        doc.add(new Paragraph(" "));

        PdfPTable tabela = new PdfPTable(5);
        tabela.setWidthPercentage(100);
        for (String h : new String[]{"Data", "Nome", "Email", "CIL", "Status"}) {
            tabela.addCell(h);
        }

        for (RelatorioEnvio r : dados) {
            tabela.addCell(r.dataEnvio().format(FMT));
            tabela.addCell(r.nome());
            tabela.addCell(r.email());
            tabela.addCell(r.cil());
            tabela.addCell(r.status());
        }
        doc.add(tabela);
        doc.close();
    }
}
