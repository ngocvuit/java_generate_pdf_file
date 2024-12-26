package com.example.demo.service;

import com.example.demo.model.PdfData;
import com.itextpdf.text.pdf.BaseFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import java.io.FileInputStream;
import java.io.FileOutputStream;

@Service
public class PdfGeneratorService {

    private static final String TEMPLATE_PATH = "D:/myFolder/template.docx"; // Template location
    private static final String OUTPUT_DIRECTORY = "D:/myFolder/";

    public void generatePdfs(List<PdfData> pdfDataList) throws Exception {
        for (PdfData data : pdfDataList) {
            // Read the Word template
            XWPFDocument doc = new XWPFDocument(new File(TEMPLATE_PATH).toURI().toURL().openStream());

            // Replace placeholders in the document
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (XWPFRun run : runs) {
                        String text = run.getText(0);
                        if (text != null) {
                            text = text.replace("{{doc_no}}", defaultString(data.getDoc_no()));
                            text = text.replace("{{cur_date}}", defaultString(data.getCur_date()));
                            text = text.replace("{{ref_full_name}}", defaultString(data.getRef_full_name()));
                            text = text.replace("{{ref_dob}}", defaultString(data.getRef_dob()));
                            text = text.replace("{{ref_nation}}", defaultString(data.getRef_nation()));
                            text = text.replace("{{ref_id_no}}", defaultString(data.getRef_id_no()));
                            text = text.replace("{{ref_id_place_iss}}", defaultString(data.getRef_id_place_iss()));
                            text = text.replace("{{ref_iss_dt}}", defaultString(data.getRef_iss_dt()));
                            text = text.replace("{{st_full_name}}", defaultString(data.getSt_full_name()));
                            text = text.replace("{{st_dob}}", defaultString(data.getSt_dob()));
                            text = text.replace("{{st_nation}}", defaultString(data.getSt_nation()));
                            text = text.replace("{{st_id_no}}", defaultString(data.getSt_id_no()));
                            text = text.replace("{{st_id_place_iss}}", defaultString(data.getSt_id_place_iss()));
                            text = text.replace("{{st_iss_dt}}", defaultString(data.getSt_iss_dt()));
                            text = text.replace("{{course}}", defaultString(data.getCourse()));
                            text = text.replace("{{total_hours}}", defaultString(data.getTotal_hours()));
                            text = text.replace("{{exp_open_dt}}", defaultString(data.getExp_open_dt()));
                            text = text.replace("{{in_score}}", defaultString(data.getIn_score()));
                            text = text.replace("{{out_score}}", defaultString(data.getIn_score()));
                            run.setText(text, 0);
                        }
                    }
                }
            }

            // Save the modified Word file to a PDF
            String outputFilePath = OUTPUT_DIRECTORY + data.getSt_full_name() + ".pdf";
            try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                // Use a library like iText or LibreOffice for Word to PDF conversion
                doc.write(out);
            }

            doc.close();
        }
    }
    // Helper method to handle null values
    private String defaultString(String value) {
        return value == null ? "" : value;
    }

    private String defaultString(String value, String defaultValue) {
        return value == null ? defaultValue : value;
    }

    public void processExcelAndGeneratePdfs(MultipartFile excelFile) throws Exception {
        // Read Excel data
        List<Map<String, String>> excelData = readExcelFile(excelFile);

        // Process each record
        generateMultiplePdfs(TEMPLATE_PATH, excelData);
    }

    private List<Map<String, String>> readExcelFile(MultipartFile excelFile) throws Exception {
        List<Map<String, String>> data = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(excelFile.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        // Get headers
        Row headerRow = sheet.getRow(0);
        List<String> headers = new ArrayList<>();
        for (Cell cell : headerRow) {
            headers.add(cell.getStringCellValue());
        }

        // Get data rows
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Map<String, String> rowData = new HashMap<>();
            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.getCell(j);
                rowData.put(headers.get(j), cell != null ? cell.toString() : "");
            }
            data.add(rowData);
        }
        workbook.close();
        return data;
    }

    private XWPFDocument processWordTemplate(Map<String, String> data, String templatePath) throws IOException {
        // Load the Word template
        try (FileInputStream fis = new FileInputStream(templatePath)) {
            XWPFDocument doc = new XWPFDocument(fis);

            // Replace placeholders in paragraphs
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : data.entrySet()) {
                            String placeholder = "{{" + entry.getKey() + "}}";
                            if (text.contains(placeholder)) {
                                text = text.replace(placeholder, entry.getValue());
                            }
                        }
                        run.setText(text, 0); // Replace the text in the run
                    }
                }
            }

            // Replace placeholders in tables (if the Word template contains tables)
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                String text = run.getText(0);
                                if (text != null) {
                                    for (Map.Entry<String, String> entry : data.entrySet()) {
                                        String placeholder = "{{" + entry.getKey() + "}}";
                                        if (text.contains(placeholder)) {
                                            text = text.replace(placeholder, entry.getValue());
                                        }
                                    }
                                    run.setText(text, 0); // Replace the text in the cell
                                }
                            }
                        }
                    }
                }
            }

            fis.close();
            return doc;
        }
    }

    private void generatePdfFromWordTemplate(Map<String, String> data, String templatePath, String outputPdfPath) throws Exception {
        // Process the Word template
        XWPFDocument doc = processWordTemplate(data, templatePath);

        // Convert the Word document to a PDF
        try (PDDocument pdf = new PDDocument()) {
            PDPage page = new PDPage();
            pdf.addPage(page);

            // Load a Unicode font using BaseFont
            String fontPath = "D:/myFolder/font/NotoSans.ttf";
            if (!new File(fontPath).exists()) {
                throw new FileNotFoundException("Font file not found: " + fontPath);
            }
            BaseFont baseFont = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

            // Create a font for PDFBox
            PDType0Font pdfFont = PDType0Font.load(pdf, new File(fontPath));

            // Start writing content to the PDF
            try (PDPageContentStream contentStream = new PDPageContentStream(pdf, page)) {
                contentStream.setFont(pdfFont, 12);
                contentStream.beginText();
                contentStream.newLineAtOffset(50, 750); // Start writing at a specific position

                // Write content from the processed Word document
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    String paragraphText = paragraph.getText();
                    if (paragraphText != null && !paragraphText.isEmpty()) {
                        contentStream.showText(paragraphText);
                        contentStream.newLineAtOffset(0, -15); // Move to the next line
                    }
                }
                contentStream.endText();
            }
            // Save the PDF
            pdf.save(outputPdfPath);
        }

        // Close the Word document
        doc.close();
    }

    public void generateMultiplePdfs(String templatePath, List<Map<String, String>> records) {
        String outputDir = "D:/GeneratedPDFs/";

        for (Map<String, String> record : records) {
            try {
                String fileName = record.get("st_full_name") + ".pdf"; // Use a unique identifier for each file
                String outputPdfPath = outputDir + fileName;
                generatePdfFromWordTemplate(record, templatePath, outputPdfPath);
                System.out.println("Generated PDF: " + outputPdfPath);
            } catch (Exception e) {
                System.err.println("Error generating PDF for record: " + record);
                e.printStackTrace();
            }
        }
    }
}

