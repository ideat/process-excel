package com.nexion.process.backend.enums;

public enum Months {
    ENERO("01"),
    FEBRERO("02"),
    MARZO("03"),
    ABRIL("04"),
    MAYO("05"),
    JUNIO("06"),
    JULIO("07"),
    AGOSTO("08"),
    SEPTIEMBRE("09"),
    OCTUBRE("10"),
    NOVIEMBRE("11"),
    DICIEMBRE("12");

    private String order;
    private Months(String order){
        this.order = order;
    }



    public String getOrder(){
        return order;
    }
}
