package com.example.demo.model;

import lombok.Data;

@Data
public class PdfData {
    private String doc_no;
    private String cur_date;
    private String ref_full_name;
    private String ref_dob;
    private String ref_nation;
    private String ref_id_no;
    private String ref_id_place_iss;
    private String ref_iss_dt;
    private String st_full_name;
    private String st_dob;
    private String st_nation;
    private String st_id_no;
    private String st_id_place_iss;
    private String st_iss_dt;
    private String course;
    private String total_hours;
    private String exp_open_dt;
    private String in_score;
    private String out_score;
}
