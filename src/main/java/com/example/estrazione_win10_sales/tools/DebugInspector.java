package com.example.estrazione_win10_sales.tools;

import com.example.estrazione_win10_sales.model.CsvRecord;
import com.example.estrazione_win10_sales.service.TerminalColumnService;
import com.example.estrazione_win10_sales.util.ExcelUtils;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.*;
import java.util.stream.Collectors;

public class DebugInspector {
    public static void main(String[] args) throws Exception {
        Path csvPath = Path.of("src/main/resources/inputfile/estrazione.csv");
        List<String[]> rows;
        try (FileReader fr = new FileReader(csvPath.toFile())) {
            CSVReader reader = new CSVReaderBuilder(fr)
                    .withCSVParser(new CSVParserBuilder().withSeparator(';').build())
                    .build();
            rows = reader.readAll();
        }
        System.out.println("CSV rows: " + rows.size());
        List<CsvRecord> records = new ArrayList<>();
        for (int i = 1; i < rows.size(); i++) {
            try {
                records.add(CsvRecord.fromCsvRow(rows.get(i)));
            } catch (Exception ex) {
                // ignore
            }
        }
        System.out.println("Parsed CsvRecord count: " + records.size());
        Set<String> distinctRice = records.stream().map(CsvRecord::getRiceId).filter(Objects::nonNull).collect(Collectors.toSet());
        System.out.println("Distinct riceIds in CSV: " + distinctRice.size());
        System.out.println("Sample riceIds from CSV: " + distinctRice.stream().limit(10).collect(Collectors.toList()));

        TerminalColumnService svc = new TerminalColumnService();
        Map<String, List<String>> terminals = svc.computeTerminalsByRiceId(records);
        System.out.println("terminalsByRiceId size: " + terminals.size());
        System.out.println("Sample terminals entries: ");
        int c = 0;
        for (Map.Entry<String, List<String>> e : terminals.entrySet()) {
            System.out.println(e.getKey() + " -> " + e.getValue());
            c++; if (c>=10) break;
        }

        // read template
        Path templatePath = Path.of("src/main/resources/inputfile/template.xlsx");
        try (InputStream in = new FileInputStream(templatePath.toFile()); Workbook wb = ExcelUtils.openWorkbook(in)) {
            Sheet s = wb.getSheetAt(0);
            Map<String,Integer> sheetRice = new LinkedHashMap<>();
            for (int r = 1; r <= s.getLastRowNum(); r++) {
                Row row = s.getRow(r);
                if (row == null) continue;
                Cell c0 = row.getCell(0);
                String v = ExcelUtils.getStringCellValue(c0);
                if (v != null && !v.isBlank()) sheetRice.put(v, r);
            }
            System.out.println("Template riceIds count: " + sheetRice.size());
            System.out.println("Sample template riceIds: " + sheetRice.keySet().stream().limit(10).collect(Collectors.toList()));

            // find intersection example
            Set<String> intersection = new LinkedHashSet<>();
            for (String k : terminals.keySet()) {
                if (sheetRice.containsKey(k)) intersection.add(k);
            }
            System.out.println("Direct intersection size: " + intersection.size());
            System.out.println("Direct intersection sample: " + intersection.stream().limit(10).collect(Collectors.toList()));

            // try numeric-only intersection
            Set<String> numericMatch = new LinkedHashSet<>();
            for (String k : terminals.keySet()) {
                String kd = k.replaceAll("\\D+", "");
                for (String t : sheetRice.keySet()) {
                    String td = t.replaceAll("\\D+", "");
                    if (!kd.isEmpty() && kd.equals(td)) {
                        numericMatch.add(k);
                        break;
                    }
                }
            }
            System.out.println("Numeric-only match size: " + numericMatch.size());
            System.out.println("Numeric-only sample: " + numericMatch.stream().limit(10).collect(Collectors.toList()));
        }
    }
}
