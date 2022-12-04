package com.pcalouche.excelspringboot.excel;

import com.pcalouche.excelspringboot.excel.data.DataValidationCustom;
import com.pcalouche.excelspringboot.excel.data.ReportData;
import com.pcalouche.excelspringboot.util.DownloadableFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.util.DefaultTempFileCreationStrategy;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.MediaType;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicReference;
import java.util.stream.Collectors;

public abstract class ExcelExport implements DownloadableFile {
    // Currently Apache POI doesn't use the new Java Time for its methods, so falling back to old Java classes
    private static final Logger logger = LoggerFactory.getLogger(ExcelExport.class);
    private static final Path EXCEL_TEMP_FILE_PATH;

    // In the static block of this class, set the Apache POI temp files path and delete an old files that may have been there.
    // Because this is a static block it should only run once when any instance of ExcelExport is created
    static {
        EXCEL_TEMP_FILE_PATH = Paths.get(System.getProperty("user.dir")).resolve("excelStaging");
        logger.info("Excel temp file staging path set to->" + EXCEL_TEMP_FILE_PATH.toAbsolutePath());
        if (!Files.exists(EXCEL_TEMP_FILE_PATH)) {
            try {
                Files.createDirectories(EXCEL_TEMP_FILE_PATH);
            } catch (IOException e) {
                logger.error("Could not create temp page", e);
            }
        }
        TempFile.setTempFileCreationStrategy(new DefaultTempFileCreationStrategy(EXCEL_TEMP_FILE_PATH.toFile()));
        // Deleting an old files at startup
        try {
            for (Path path : Files.list(EXCEL_TEMP_FILE_PATH).collect(Collectors.toList())) {
                logger.info("Deleting old Excel temp file->" + path.toAbsolutePath());
                Files.deleteIfExists(path);
            }
        } catch (IOException e) {
            logger.error("Exception occurred during Excel temp file deletion", e);
        }
    }

    protected final Font boldFont;
    protected final XSSFCellStyle boldStyle;
    private final Workbook workbook;

    public ExcelExport(boolean useStreaming) {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
        boldFont = xssfWorkbook.createFont();
        boldFont.setBold(true);
        boldStyle = xssfWorkbook.createCellStyle();
        boldStyle.setFont(boldFont);

        // If streaming is true use the streaming version of Apache Poi.
        if (useStreaming) {
            workbook = new SXSSFWorkbook(xssfWorkbook, 100);
        } else {
            workbook = xssfWorkbook;
        }
    }

    @Override
    public byte[] getBytes() {
        Long startTime = System.currentTimeMillis();
        logger.info("start getBytes for Excel");
        byte[] bytes;
        try (ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream()) {
            workbook.write(byteArrayOutputStream);
            bytes = byteArrayOutputStream.toByteArray();
            // If streaming Apache Poi was used then dispose of temp files
            if (workbook instanceof SXSSFWorkbook) {
                ((SXSSFWorkbook) workbook).dispose();
            }
        } catch (IOException e) {
            logger.error("Failed to convert Excel to byte array", e);
            throw new RuntimeException("Failed to convert Excel to byte array: " + e.getMessage());
        }
        Long endTime = System.currentTimeMillis();
        logger.info("end getBytes for Excel->" + (endTime - startTime) / 1000 + " seconds.");
        return bytes;
    }

    @Override
    public MediaType getContentType() {
        return MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    }

    protected XSSFWorkbook getXSSFWorkbook() {
        return (XSSFWorkbook) workbook;
    }

    protected SXSSFWorkbook getSXSSFWorkbook() {
        return (SXSSFWorkbook) workbook;
    }

    protected void addDataToWorkbook(final Workbook workbook,
                                     final ReportData reportData) {
        Long sheetStartTime = System.currentTimeMillis();
        reportData.build();
        logger.info("start sheet generation");
        SXSSFSheet dataSheet = (SXSSFSheet) workbook.createSheet("Data");
        AtomicReference<Integer> count = new AtomicReference<>(0);
        createHeader(reportData.getTitles(), count, dataSheet);
        while (reportData.hasNext()) {
            createRow(reportData.next(), count, dataSheet);
        }

        Optional.ofNullable(
                reportData.getValidations())
                        .map(dataValidationCustoms ->
                            addDropdowns(dataSheet, dataValidationCustoms, count)
                        );

        Long endSheetTime = System.currentTimeMillis();
        logger.info("Total sheet generation time " + (endSheetTime - sheetStartTime) / 1000 + " seconds.");
    }

    private void createHeader(final List<String> objectList,
                           final AtomicReference<Integer> count,
                           final SXSSFSheet dataSheet) {

        AtomicReference<Integer> countColum = new AtomicReference<>(0);
        Row row = dataSheet.createRow(count.get());

        objectList.forEach(value -> {
            Cell cell = row.createCell(countColum.get());
            cell.setCellValue(value);
            CellStyle styleTitle = workbook.createCellStyle();
            Font font = workbook.createFont();
            workbook.createCellStyle();
            font.setBold(Boolean.TRUE);
            styleTitle.setFont(font);
            styleTitle.setWrapText(true);
            cell.setCellStyle(styleTitle);
            countColum.getAndSet(countColum.get() + 1);
        });

        count.getAndSet(count.get() +1);
    }

    private void createRow(final List<Object> objectList,
                           final AtomicReference<Integer> count,
                           final SXSSFSheet dataSheet) {

        AtomicReference<Integer> countColum = new AtomicReference<>(0);
        Row row = dataSheet.createRow(count.get());

        objectList.forEach(value -> {
            createCell(row, countColum, value);
        });

        count.getAndSet(count.get() +1);
    }

    private void createCell(Row row, AtomicReference<Integer> countColum, Object value) {
        Cell cell = row.createCell(countColum.get());
        if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        }
        countColum.getAndSet(countColum.get() + 1);
    }

    private SXSSFSheet addDropdowns(final SXSSFSheet dataSheet,
                              final List<DataValidationCustom> dataValidationCustom,
                              final AtomicReference<Integer> count) {

        dataValidationCustom.forEach(dataValidationCustom1 -> {
            DataValidationHelper validationHelper = dataSheet.getDataValidationHelper();
            CellRangeAddressList addressList =
                    new CellRangeAddressList(1, count.get() - 1, dataValidationCustom1.getColumn(), dataValidationCustom1.getColumn());
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(dataValidationCustom1.getOptions());
            DataValidation dataValidation  = validationHelper.createValidation(constraint, addressList);
            dataValidation.setSuppressDropDownArrow(true);
            dataSheet.addValidationData(dataValidation);
        });

        return dataSheet;
    }




}