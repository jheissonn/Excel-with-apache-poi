package com.pcalouche.excelspringboot.excel;

import com.pcalouche.excelspringboot.excel.data.ReportDataExampleImpl;

public class StreamExcelExport extends ExcelExport {

    public StreamExcelExport() {
        super(true);
        build();
    }

    private void build() {
        addDataToWorkbook(getSXSSFWorkbook(), new ReportDataExampleImpl());
    }

    @Override
    public String getFilename() {
        return "streaming.xlsx";
    }
}
