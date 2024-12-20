package com.example.demo.controller;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/api")
public class HomeController {

    @CrossOrigin(origins = "http://127.0.0.1:5500") // Allow requests from this origin
    @PostMapping("/generate-pdfs/home")
    public ResponseEntity<byte[]> generatePDFs(@RequestParam("jsonFile") MultipartFile jsonFile) {
        try {
            // Step 1: Parse JSON File
            ObjectMapper objectMapper = new ObjectMapper();
            JsonNode jsonData = objectMapper.readTree(jsonFile.getInputStream());

            // Step 2: Prepare the Word Template
            String templatePath = "src/main/resources/template.docx";
            XWPFDocument template = new XWPFDocument(new FileInputStream(templatePath));

            // Step 3: Prepare ZIP Output
            ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
            ZipOutputStream zipOut = new ZipOutputStream(zipOutputStream);

            // Step 4: Generate PDFs for Each Record
            int fileIndex = 1;
            for (JsonNode record : jsonData) {
                // Extract dynamic data
                Map<String, String> placeholders = new HashMap<>();
                record.fields().forEachRemaining(field -> placeholders.put(field.getKey(), field.getValue().asText()));

                // Populate the Word Template
                XWPFDocument populatedDoc = replacePlaceholders(template, placeholders);

                // Convert to PDF (implement your conversion logic)
                ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
                convertWordToPdf(populatedDoc, pdfOutputStream);

                // Add PDF to ZIP
                ZipEntry zipEntry = new ZipEntry("file_" + fileIndex++ + ".pdf");
                zipOut.putNextEntry(zipEntry);
                zipOut.write(pdfOutputStream.toByteArray());
                zipOut.closeEntry();
            }

            zipOut.close();

            // Step 5: Return ZIP File
            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"generated_pdfs.zip\"")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(zipOutputStream.toByteArray());

        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.internalServerError().body(null);
        }
    }

    private XWPFDocument replacePlaceholders(XWPFDocument doc, Map<String, String> placeholders) {
        for (XWPFParagraph paragraph : doc.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : placeholders.entrySet()) {
                        text = text.replace("{{" + entry.getKey() + "}}", entry.getValue());
                    }
                    run.setText(text, 0);
                }
            }
        }
        return doc;
    }

    private void convertWordToPdf(XWPFDocument doc, OutputStream outputStream) {
        // Implement your Word-to-PDF conversion logic here.
        // Example: Use libraries like PDFBox, iText, or third-party services.
    }
}


