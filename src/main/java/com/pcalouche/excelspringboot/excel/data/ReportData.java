package com.pcalouche.excelspringboot.excel.data;

import java.util.List;

public interface ReportData {

    void build();

    boolean hasNext();

    List<Object> next();

    List<String> getTitles();

    default List<DataValidationCustom> getValidations() {
        return null;
    }



}
