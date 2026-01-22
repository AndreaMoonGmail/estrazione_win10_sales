package com.example.estrazione_win10_sales.tools;

import com.example.estrazione_win10_sales.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

public class OutputInspector {
    public static void main(String[] args) throws Exception {
        Path p = Path.of("C:/Users/sartirana/Documents/workspace_intellij_estrazione_win10_sales/target/estrazione_output.xlsx");
        try (InputStream in = new FileInputStream(p.toFile()); Workbook wb = ExcelUtils.openWorkbook(in)) {
            Sheet s = wb.getSheetAt(0);
            Map<String,Integer> headers = ExcelUtils.headerMap(s);
            System.out.println("Headers: " + headers);
            Map<String,Integer> sheetRice = new LinkedHashMap<>();
            for (int r = 1; r <= s.getLastRowNum(); r++) {
                Row row = s.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                String v = ExcelUtils.getStringCellValue(c0);
                if (v != null && !v.isBlank()) sheetRice.put(v, r);
            }
            String target = "285353";
            if (!sheetRice.containsKey(target)) {
                System.out.println("RiceId not found in sheet: " + target);
                return;
            }
            int rowIndex = sheetRice.get(target);
            Row row = s.getRow(rowIndex);
            System.out.println("Row index for " + target + " = " + rowIndex);
            for (Map.Entry<String,Integer> e : headers.entrySet()) {
                String h = e.getKey();
                int ci = e.getValue();
                Cell c = row.getCell(ci);
                String val = ExcelUtils.getStringCellValue(c);
                System.out.println(h + " -> '" + val + "'");
            }
        }
    }
}
