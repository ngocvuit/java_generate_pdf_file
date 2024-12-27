package com.example.demo.util;

import com.example.demo.model.PdfData;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.text.Normalizer;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class CommonUtils {
    public static String generateFileName(String name) {
        if (name == null || name.isEmpty()) {
            return "";
        }
        String withoutAccents = Normalizer.normalize(name, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
        return withoutAccents.replaceAll("\\s+", "_").toUpperCase();
    }

    // Helper method to handle null values
    public static String defaultString(String value) {
        return value == null ? "" : value;
    }

    public static void replacePlaceholder(Map<String, String> data, List<XWPFParagraph> paragraphs) {
        fillDataToWordFile(paragraphs, data);
    }

    public static void replacePlaceholder(PdfData data, List<XWPFParagraph> paragraphs) {
        Map<String, String> placeholderMap = new HashMap<>();
        placeholderMap.put("doc_no", data.getDoc_no());
        placeholderMap.put("cur_date", data.getCur_date());
        placeholderMap.put("ref_full_name", data.getRef_full_name());
        placeholderMap.put("ref_dob", data.getRef_dob());
        placeholderMap.put("ref_nation", data.getRef_nation());
        placeholderMap.put("ref_id_no", data.getRef_id_no());
        placeholderMap.put("ref_id_place_iss", data.getRef_id_place_iss());
        placeholderMap.put("ref_iss_dt", data.getRef_iss_dt());
        placeholderMap.put("st_full_name", data.getSt_full_name());
        placeholderMap.put("st_dob", data.getSt_dob());
        placeholderMap.put("st_nation", data.getSt_nation());
        placeholderMap.put("st_id_no", data.getSt_id_no());
        placeholderMap.put("st_id_place_iss", data.getSt_id_place_iss());
        placeholderMap.put("st_iss_dt", data.getSt_iss_dt());
        placeholderMap.put("course", data.getCourse());
        placeholderMap.put("total_hours", data.getTotal_hours());
        placeholderMap.put("exp_open_dt", data.getExp_open_dt());
        placeholderMap.put("in_score", data.getIn_score());
        placeholderMap.put("out_score", data.getOut_score());

        fillDataToWordFile(paragraphs, placeholderMap);
    }

    public  static Map<String, String> createRecordFromPdfData(PdfData data) {
        Map<String, String> record = new HashMap<>();
        record.put("doc_no", data.getDoc_no());
        record.put("cur_date", data.getCur_date());
        record.put("ref_full_name", data.getRef_full_name());
        record.put("ref_dob", data.getRef_dob());
        record.put("ref_nation", data.getRef_nation());
        record.put("ref_id_no", data.getRef_id_no());
        record.put("ref_id_place_iss", data.getRef_id_place_iss());
        record.put("ref_iss_dt", data.getRef_iss_dt());
        record.put("st_full_name", data.getSt_full_name());
        record.put("st_dob", data.getSt_dob());
        record.put("st_nation", data.getSt_nation());
        record.put("st_id_no", data.getSt_id_no());
        record.put("st_id_place_iss", data.getSt_id_place_iss());
        record.put("st_iss_dt", data.getSt_iss_dt());
        record.put("course", data.getCourse());
        record.put("total_hours", data.getTotal_hours());
        record.put("exp_open_dt", data.getExp_open_dt());
        record.put("in_score", data.getIn_score());
        record.put("out_score", data.getOut_score());
        return record;
    }



    private static void fillDataToWordFile(List<XWPFParagraph> paragraphs, Map<String, String> placeholderMap) {
        for (XWPFParagraph paragraph : paragraphs) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null) {
                    for (Map.Entry<String, String> entry : placeholderMap.entrySet()) {
                        String placeholder = entry.getKey();
                        if (text.contains(placeholder)) {
                            text = text.replace(placeholder, entry.getValue());
                        }
                    }
                    run.setText(text, 0); // Replace the text in the run
                }
            }
        }
    }

    public static String getStringCellValue(Row row, int cellIndex){
        Cell cell = row.getCell(cellIndex);
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                default:
                    return "";
            }
        }
        return "";
    }
}
