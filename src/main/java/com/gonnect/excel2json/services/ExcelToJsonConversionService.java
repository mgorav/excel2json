package com.gonnect.excel2json.services;

import com.gonnect.excel2json.config.ExcelToJsonConverterConfig;
import com.gonnect.excel2json.model.AWorkbook;
import com.gonnect.excel2json.model.AWorksheet;
import com.sun.media.sound.InvalidFormatException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

public class ExcelToJsonConversionService {

    private ExcelToJsonConverterConfig config = null;

    public ExcelToJsonConversionService(ExcelToJsonConverterConfig config) {
        this.config = config;
    }

    public static AWorkbook convert(ExcelToJsonConverterConfig config) throws InvalidFormatException, IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        return new ExcelToJsonConversionService(config).convert();
    }

    public AWorkbook convert()
            throws InvalidFormatException, IOException, org.apache.poi.openxml4j.exceptions.InvalidFormatException {
        AWorkbook book = new AWorkbook();

        InputStream inp = new FileInputStream(config.getSourceFile());
        Workbook wb = WorkbookFactory.create(inp);

        book.setFileName(config.getSourceFile());
        int loopLimit = wb.getNumberOfSheets();
        if (config.getNumberOfSheets() > 0 && loopLimit > config.getNumberOfSheets()) {
            loopLimit = config.getNumberOfSheets();
        }
        int rowLimit = config.getRowLimit();
        int startRowOffset = config.getRowOffset();
        int currentRowOffset = -1;
        int totalRowsAdded = 0;


        for (int i = 0; i < loopLimit; i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) {
                continue;
            }
            AWorksheet tmp = new AWorksheet();
            tmp.setName(sheet.getSheetName());
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                if (row == null) {
                    continue;
                }
                boolean hasValues = false;
                ArrayList<Object> rowData = new ArrayList<Object>();
                for (int k = 0; k <= row.getLastCellNum(); k++) {
                    Cell cell = row.getCell(k);
                    if (cell != null) {
                        Object value = cellToObject(cell);
                        hasValues = hasValues || value != null;
                        rowData.add(value);
                    } else {
                        rowData.add(null);
                    }
                }
                if (hasValues || !config.isOmitEmpty()) {
                    currentRowOffset++;
                    if (rowLimit > 0 && totalRowsAdded == rowLimit) {
                        break;
                    }
                    if (startRowOffset > 0 && currentRowOffset < startRowOffset) {
                        continue;
                    }
                    tmp.addRow(rowData);
                    totalRowsAdded++;
                }
            }
            if (config.isFillColumns()) {
                tmp.fillColumns();
            }
            book.addExcelWorksheet(tmp);
        }

        return book;
    }

    private Object cellToObject(Cell cell) {

        CellType type = cell.getCellType();

        if (type == CellType.STRING) {
            return cleanString(cell.getStringCellValue());
        }

        if (type == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        }

        if (type == CellType.NUMERIC) {

            if (cell.getCellStyle().getDataFormatString().contains("%")) {
                return cell.getNumericCellValue() * 100;
            }

            return numeric(cell);
        }

        if (type == CellType.FORMULA) {
            switch (cell.getCachedFormulaResultType()) {
                case NUMERIC:
                    return numeric(cell);
                case STRING:
                    return cleanString(cell.getRichStringCellValue().toString());
            }
        }

        return null;

    }

    private String cleanString(String str) {
        return str.replace("\n", "").replace("\r", "");
    }

    private Object numeric(Cell cell) {
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
            if (config.getFormatDate() != null) {
                return config.getFormatDate().format(cell.getDateCellValue());
            }
            return cell.getDateCellValue();
        }
        return Double.valueOf(cell.getNumericCellValue());
    }
}
