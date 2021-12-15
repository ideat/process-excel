package com.nexion.process.backend.entity;

import lombok.Data;

import java.util.List;
import java.util.UUID;

@Data
public class Configuration {
    private UUID id;

    private String nameFile;

    private String configList;

    /*
    name_sheet: ,
    column_category: ,
    row_category: ,
    col_parameter: ,
    row_parameter: ,
    number_rows_parameter:,
    row_start_date: ,
    col_start_date: ,
    col_space_date: ,
    total_col_date: ,
    type_date:, month, date, year, semester
    row_start_sample:,
    col_start_sample:,
    col_space_sample: ;
     */
}
