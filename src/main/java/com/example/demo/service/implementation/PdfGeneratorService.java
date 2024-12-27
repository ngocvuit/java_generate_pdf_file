package com.example.demo.service.implementation;

import com.example.demo.model.PdfData;
import com.example.demo.service.IExcelService;
import com.example.demo.service.IPdfGeneratorService;
import com.example.demo.service.IWordService;
import com.example.demo.util.CommonUtils;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

@Service
public class PdfGeneratorService implements IPdfGeneratorService {
    private static final String TEMPLATE_PATH = "src/main/resources/word/template.docx"; // Template location
    private static final String OUTPUT_DIRECTORY = "src/main/resources/pdf/result/";
    private static final String OUTPUT_EXTENSION = ".pdf";

    private final IWordService wordService;

    private final IExcelService excelService;

    public PdfGeneratorService(IExcelService excelService, IWordService wordService) {
        this.excelService = excelService;
        this.wordService = wordService;
    }

    @Override
    public void generatePdfs(List<PdfData> pdfDataList)  {
        for (PdfData data : pdfDataList) {
            this.generateSinglePdf(data);
        }
    }

    @Override
    public void processExcelAndGeneratePdfs(MultipartFile excelFile) throws Exception {
        List<PdfData> excelData = excelService.readExcelFilePdfData(excelFile);
        generateMultiplePdfs(excelData); //Generate PDF file for each row
    }

    private void generateMultiplePdfs(List<PdfData> pdfDataList) {
        for (PdfData data : pdfDataList) {
            convertWordToPdf(data);
        }
    }

    private void generateSinglePdf(PdfData data) {
        convertWordToPdf(data);
    }

    private void convertWordToPdf(PdfData data) {
        String fileName = CommonUtils.generateFileName(data.getSt_full_name()) + OUTPUT_EXTENSION;
        try {
            String outputPdfPath = OUTPUT_DIRECTORY + fileName;
            Map<String, String> record = CommonUtils.createRecordFromPdfData(data);
            wordService.convertWordToPDF(record, TEMPLATE_PATH, outputPdfPath);
            System.out.println("Generated PDF: " + outputPdfPath);
        } catch (Exception e) {
            System.err.println("Error generating PDF: " + fileName);
            e.printStackTrace();
        }
    }

}

