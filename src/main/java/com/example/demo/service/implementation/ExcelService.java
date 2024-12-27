package com.example.demo.service.implementation;

import com.example.demo.model.PdfData;
import com.example.demo.service.IExcelService;
import com.example.demo.util.CommonUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService implements IExcelService {
    @Override
    public List<Map<String, String>> readExcelFileMap(MultipartFile excelFile) throws Exception {
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
                rowData.put(headers.get(j).trim(), cell != null ? cell.toString().trim() : "");
            }
            data.add(rowData);
        }
        workbook.close();
        return data;
    }

    @Override
    public List<PdfData> readExcelFilePdfData(MultipartFile excelFile) throws Exception {
        List<PdfData> pdfDataList = new ArrayList<>();
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
            PdfData pdfData = new PdfData();

            for (int j = 0; j < headers.size(); j++) {
                Cell cell = row.getCell(j);
                String value = cell != null ? cell.toString().trim() : "";

                // Set the value into the PdfData object based on the header
                switch (headers.get(j).trim()) {
                    case "doc_no":
                        pdfData.setDoc_no(value);
                        break;
                    case "cur_date":
                        pdfData.setCur_date(value);
                        break;
                    case "ref_full_name":
                        pdfData.setRef_full_name(value);
                        break;
                    case "ref_dob":
                        pdfData.setRef_dob(value);
                        break;
                    case "ref_nation":
                        pdfData.setRef_nation(value);
                        break;
                    case "ref_id_no":
                        pdfData.setRef_id_no(value);
                        break;
                    case "ref_id_place_iss":
                        pdfData.setRef_id_place_iss(value);
                        break;
                    case "ref_iss_dt":
                        pdfData.setRef_iss_dt(value);
                        break;
                    case "st_full_name":
                        pdfData.setSt_full_name(value);
                        break;
                    case "st_dob":
                        pdfData.setSt_dob(value);
                        break;
                    case "st_nation":
                        pdfData.setSt_nation(value);
                        break;
                    case "st_id_no":
                        pdfData.setSt_id_no(value);
                        break;
                    case "st_id_place_iss":
                        pdfData.setSt_id_place_iss(value);
                        break;
                    case "st_iss_dt":
                        pdfData.setSt_iss_dt(value);
                        break;
                    case "course":
                        pdfData.setCourse(value);
                        break;
                    case "total_hours":
                        pdfData.setTotal_hours(value);
                        break;
                    case "exp_open_dt":
                        pdfData.setExp_open_dt(value);
                        break;
                    case "in_score":
                        pdfData.setIn_score(value);
                        break;
                    case "out_score":
                        pdfData.setOut_score(value);
                        break;
                    default:
                        // Handle unexpected columns if necessary
                        break;
                }
            }

            pdfDataList.add(pdfData);
        }

        workbook.close();
        return pdfDataList;
    }


    @Override
    public PdfData convertExcelToPdfData(MultipartFile file) throws IOException {
        PdfData pdfData = new PdfData();

        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(1);
            pdfData.setDoc_no(CommonUtils.getStringCellValue(row, 0));
            pdfData.setCur_date(CommonUtils.getStringCellValue(row, 1));
            pdfData.setRef_full_name(CommonUtils.getStringCellValue(row, 2));
            pdfData.setRef_dob(CommonUtils.getStringCellValue(row, 3));
            pdfData.setRef_nation(CommonUtils.getStringCellValue(row, 4));
            pdfData.setRef_id_no(CommonUtils.getStringCellValue(row, 5));
            pdfData.setRef_id_place_iss(CommonUtils.getStringCellValue(row, 6));
            pdfData.setRef_iss_dt(CommonUtils.getStringCellValue(row, 7));
            pdfData.setSt_full_name(CommonUtils.getStringCellValue(row, 8));
            pdfData.setSt_dob(CommonUtils.getStringCellValue(row, 9));
            pdfData.setSt_nation(CommonUtils.getStringCellValue(row, 10));
            pdfData.setSt_id_no(CommonUtils.getStringCellValue(row, 11));
            pdfData.setSt_id_place_iss(CommonUtils.getStringCellValue(row, 12));
            pdfData.setSt_iss_dt(CommonUtils.getStringCellValue(row, 13));
            pdfData.setCourse(CommonUtils.getStringCellValue(row, 14));
            pdfData.setTotal_hours(CommonUtils.getStringCellValue(row, 15));
            pdfData.setExp_open_dt(CommonUtils.getStringCellValue(row, 16));
            pdfData.setIn_score(CommonUtils.getStringCellValue(row, 17));
            pdfData.setOut_score(CommonUtils.getStringCellValue(row, 18));

        } catch (IOException e) {
            throw new IOException("Error processing the Excel file: " + e.getMessage(), e);
        }

        return pdfData;
    }

}
