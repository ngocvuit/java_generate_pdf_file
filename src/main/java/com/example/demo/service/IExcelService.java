package com.example.demo.service;

import com.example.demo.model.PdfData;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface IExcelService {
    List<Map<String, String>> readExcelFileMap(MultipartFile excelFile) throws Exception;
    List<PdfData> readExcelFilePdfData(MultipartFile excelFile) throws Exception;
    PdfData convertExcelToPdfData(MultipartFile file) throws IOException;
}
