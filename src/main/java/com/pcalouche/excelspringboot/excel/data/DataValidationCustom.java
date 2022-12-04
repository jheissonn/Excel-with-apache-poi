package com.pcalouche.excelspringboot.excel.data;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@AllArgsConstructor
public class DataValidationCustom {
    private Integer column;
    private String[] options;
}
