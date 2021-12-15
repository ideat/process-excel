package com.nexion.process.backend.entity;

import lombok.Data;

@Data
public class CellData {

    private Integer rowInit;
    private Integer colInit;
    private Integer rowEnd;
    private Integer colEnd;
    private Integer numSkipColumn;

}
