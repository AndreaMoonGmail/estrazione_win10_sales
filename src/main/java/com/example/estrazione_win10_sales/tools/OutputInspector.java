package com.example.estrazione_win10_sales.tools;

import com.example.estrazione_win10_sales.util.ExcelUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.Map;

public class OutputInspector {
    public static void main(String[] args) throws Exception {
        Path p = Path.of("C:/Users/sartirana/Documents/workspace_intellij_estrazione_win10_sales/target/estrazione_output.xlsx");
        Path out = Path.of("inspect_285353.txt");
        StringBuilder sb = new StringBuilder();
        try (InputStream in = new FileInputStream(p.toFile()); Workbook wb = ExcelUtils.openWorkbook(in)) {
            Sheet s = wb.getSheetAt(0);
            Map<String,Integer> headers = ExcelUtils.headerMap(s);
            sb.append("Headers: ").append(headers).append(System.lineSeparator());
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
                sb.append("RiceId not found in sheet: ").append(target).append(System.lineSeparator());
                Files.writeString(out, sb.toString());
                return;
            }
            int rowIndex = sheetRice.get(target);
            Row row = s.getRow(rowIndex);
            sb.append("Row index for ").append(target).append(" = ").append(rowIndex).append(System.lineSeparator());
            for (Map.Entry<String,Integer> e : headers.entrySet()) {
                String h = e.getKey();
                int ci = e.getValue();
                Cell c = row.getCell(ci);
                String val = ExcelUtils.getStringCellValue(c);
                sb.append(h).append(" -> '").append(val).append("'\n");
            }
            Files.writeString(out, sb.toString());
        }
    }
}
