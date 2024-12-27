package com.example.demo.service;

import com.example.demo.model.PdfData;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface IPdfGeneratorService {
    void generatePdfs(List<PdfData> pdfDataList) throws Exception ;
    void processExcelAndGeneratePdfs(MultipartFile excelFile) throws Exception ;
}
