package com.pcalouche.excelspringboot.excel;

import com.pcalouche.excelspringboot.excel.data.ReportDataExampleImpl;
import com.pcalouche.excelspringboot.util.DownloadableFile;

public class NonStreamingExcelExport extends ExcelExport implements DownloadableFile {

    public NonStreamingExcelExport() {
        super(false);
    }

    private void build() {
        addDataToWorkbook(getXSSFWorkbook(), new ReportDataExampleImpl());
    }

    @Override
    public String getFilename() {
        return "non-streaming.xlsx";
    }
}