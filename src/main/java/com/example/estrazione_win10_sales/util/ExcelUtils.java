package com.example.estrazione_win10_sales.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;

public final class ExcelUtils {

    private ExcelUtils() {}

    public static Workbook openWorkbook(InputStream in) throws IOException {
        return new XSSFWorkbook(in);
    }

    /**
     * Returns a map of header name -> column index for the given header row (first non-empty row)
     */
    public static Map<String, Integer> headerMap(Sheet sheet) {
        Map<String, Integer> map = new LinkedHashMap<>();
        Row header = sheet.getRow(0);
        if (header == null) return map;
        for (Cell c : header) {
            String value = getStringCellValue(c);
            if (value != null && !value.isBlank()) {
                map.put(value, c.getColumnIndex());
            }
        }
        return map;
    }

    public static String getStringCellValue(Cell c) {
        if (c == null) return null;
        if (c.getCellType() == CellType.STRING) return c.getStringCellValue();
        if (c.getCellType() == CellType.NUMERIC) {
            double d = c.getNumericCellValue();
            if (Double.isFinite(d) && d == Math.floor(d)) {
                // whole number, format without decimal
                long l = (long) d;
                return Long.toString(l);
            }
            // otherwise preserve decimal representation (avoid scientific notation)
            DecimalFormat df = new DecimalFormat("0.#####");
            return df.format(d);
        }
        if (c.getCellType() == CellType.BOOLEAN) return String.valueOf(c.getBooleanCellValue());
        if (c.getCellType() == CellType.FORMULA) return c.getCellFormula();
        return c.toString();
    }

    public static Cell createCellIfAbsent(Row row, int colIndex) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);
        return cell;
    }

}
