package com.pcalouche.excelspringboot.excel.data;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Setter
@Getter
@NoArgsConstructor
@AllArgsConstructor
public class ExampleData {
    private Integer fieldExampleInteger;
    private String fieldExampleString;
    private String fieldExampleDropDown;
    private Double fieldExampleValue;
}
