package com.example.estrazione_win10_sales.tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;

public class ExcelInspector {
    public static void main(String[] args) throws Exception {
        String path = "target/estrazione_output.xlsx";
        try (InputStream in = new FileInputStream(path); Workbook wb = new XSSFWorkbook(in)) {
            Sheet s = wb.getSheetAt(0);
            Row header = s.getRow(0);
            System.out.println("Header values:");
            if (header != null) {
                for (Cell c : header) {
                    System.out.print((c.getStringCellValue()) + " | ");
                }
                System.out.println();
            }
            int rows = s.getPhysicalNumberOfRows();
            System.out.println("Row count: " + rows);
            int toPrint = Math.min(5, rows-1);
            for (int i = 1; i <= toPrint; i++) {
                Row r = s.getRow(i);
                if (r == null) continue;
                for (Cell c : r) {
                    System.out.print(ExcelInspector.cellToString(c) + " | ");
                }
                System.out.println();
            }
        }
    }

    private static String cellToString(Cell c) {
        if (c == null) return "";
        switch (c.getCellType()) {
            case STRING: return c.getStringCellValue();
            case NUMERIC: return String.valueOf(c.getNumericCellValue());
            case BOOLEAN: return String.valueOf(c.getBooleanCellValue());
            default: return c.toString();
        }
    }
}
