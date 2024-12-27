package com.example.demo.service;

import java.util.Map;

public interface IWordService {
    void convertWordToPDF(Map<String, String> data, String templatePath, String outputPdfPath) throws Exception;

}
