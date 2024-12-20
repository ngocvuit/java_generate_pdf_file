package com.example.demo.service;

import com.example.demo.model.PdfData;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

@Service
public class PdfGeneratorService {

    private static final String TEMPLATE_PATH = "src/main/resources/template.docx"; // Template location
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
}

