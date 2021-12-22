package com.nexion.process.backend.entity;

import lombok.Data;

@Data
public class ConfigSheet {

    private Integer idEpsa;
    private String nameSheet;
    private String cellCategory;
    private String staticCategory;


    private String cellParameter;
    private Integer numberRowsParameter;

    private String cellStartDate;
    private Integer colSpaceDate ;
    private Integer totalColDate ;
    private String typeDate; //month; date; year; semester; process; static
    private String staticDate;
    private String conditionProcess;
    private String year;
    private String day;

    private String cellStartSampleName;
    private Integer colSpaceSampleName;
    private String typeSample; //process
    private String staticSampleName;
    private String cellSampleName;

    private String cellValueSample;

//    private String cellValueUnit;

    private String cellPointSample;
    private String colSpacePointSample;
    private String staticPointSample; //true, false

    private String cellSamplingLocation;
    private String staticSamplingLocation;// true, false

    private String cellUnitMeasure;

}
