package com.example.demo.controller;

import com.example.demo.model.PdfData;
import com.example.demo.service.IPdfGeneratorService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@RequestMapping("/api")
public class PdfController {

    @Autowired
    private IPdfGeneratorService pdfGeneratorService;

    @PostMapping("/generate-pdfs")
    public ResponseEntity<String> generatePdfs(@RequestBody List<PdfData> pdfDataList) {
        try {
            if (pdfDataList == null || pdfDataList.isEmpty()) {
                return ResponseEntity.badRequest().body("Input data is empty.");
            }
            pdfGeneratorService.generatePdfs(pdfDataList);
            return ResponseEntity.ok("PDFs generated successfully!");
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body("Failed to generate PDFs.");
        }
    }

    @PostMapping({"/upload-excel"})
    public ResponseEntity<String> uploadExcelAndGeneratePdf(@RequestParam("file") MultipartFile file) {
        try {
            pdfGeneratorService.processExcelAndGeneratePdfs(file);
            return ResponseEntity.ok("PDF files generated successfully.");
        } catch (Exception e) {
            e.printStackTrace();
            return ResponseEntity.status(500).body("Failed to generate PDFs: " + e.getMessage());
        }
    }
}


