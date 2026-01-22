package com.example.estrazione_win10_sales.tools;

import com.example.estrazione_win10_sales.model.CsvRecord;
import com.example.estrazione_win10_sales.service.TerminalColumnService;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;

import java.io.FileReader;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class TerminalDebug {
    public static void main(String[] args) throws Exception {
        Path csvPath = Path.of("src/main/resources/inputfile/estrazione.csv");
        List<String[]> rows;
        try (FileReader fr = new FileReader(csvPath.toFile());
             CSVReader reader = new CSVReaderBuilder(fr)
                    .withCSVParser(new CSVParserBuilder().withSeparator(';').build())
                    .build()) {
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

        TerminalColumnService svc = new TerminalColumnService();
        Map<String, java.util.List<String>> terminals = svc.computeTerminalsByRiceId(records);

        Path out = Path.of("terminal_debug_out.txt");
        StringBuilder sb = new StringBuilder();
        sb.append("terminalsByRiceId size: ").append(terminals.size()).append(System.lineSeparator());
        String[] targets = new String[] {"283840","285353"};
        for (String target : targets) {
            if (terminals.containsKey(target)) {
                sb.append(target).append(" -> ").append(terminals.get(target)).append(System.lineSeparator());
            } else {
                sb.append("No entry for ").append(target).append(System.lineSeparator());
            }
        }
        java.nio.file.Files.writeString(out, sb.toString());
    }
}
