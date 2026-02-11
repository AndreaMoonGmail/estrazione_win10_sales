package com.example.estrazione_win10_sales.service;

import com.example.estrazione_win10_sales.model.CsvRecord;
import com.example.estrazione_win10_sales.util.ExcelUtils;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class CsvToExcelService {

    /**
     * Run conversion: read CSV, add columns to template as needed, write output file.
     * CSV is expected to have semicolon (;) as separator in this project.
     */
    public void run(String csvPath, String templatePath, String outputPath) throws IOException {
        // Resolve resources: support classpath: prefix
        InputStream templateIn = openResource(templatePath);
        if (templateIn == null) {
            throw new FileNotFoundException("Template not found: " + templatePath);
        }

        try (Workbook workbook = new XSSFWorkbook(templateIn)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, Integer> existing = ExcelUtils.headerMap(sheet);

            // Read CSV headers and rows
            List<String[]> rows = readCsv(csvPath);
            if (rows.isEmpty()) {
                // no-op
                return;
            }

            String[] headers = rows.get(0);

            // Convert CSV rows into CsvRecord list (skip header row)
            List<CsvRecord> records = new ArrayList<>();
            for (int i = 1; i < rows.size(); i++) {
                try {
                    records.add(CsvRecord.fromCsvRow(rows.get(i)));
                } catch (IllegalArgumentException ex) {
                    // skip malformed row
                }
            }

            // Compute terminal columns per riceId
            TerminalColumnService terminalService = new TerminalColumnService();
            Map<String, List<String>> terminalsByRiceId = terminalService.computeTerminalsByRiceId(records);
            int maxTerminals = terminalService.maxTerminalCount(terminalsByRiceId);

            // Ensure terminal headers exist: terminal_1 .. terminal_N
            List<String> terminalHeaders = new ArrayList<>();
            for (int t = 1; t <= maxTerminals; t++) {
                terminalHeaders.add("terminal_" + t);
            }

            // determine which headers from CSV are new (not in existing map)
            List<String> newHeaders = new ArrayList<>(Arrays.stream(headers)
                    .filter(h -> h != null && !h.isBlank())
                    .filter(h -> !existing.containsKey(h))
                    .collect(Collectors.toList()));

            // append terminal headers too if missing
            for (String th : terminalHeaders) {
                if (!existing.containsKey(th)) newHeaders.add(th);
            }

            // add new headers at the end
            int lastCol = existing.isEmpty() ? 0 : existing.values().stream().max(Integer::compareTo).orElse(-1) + 1;
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) headerRow = sheet.createRow(0);
            for (int i = 0; i < newHeaders.size(); i++) {
                Cell c = ExcelUtils.createCellIfAbsent(headerRow, lastCol + i);
                c.setCellValue(newHeaders.get(i));
                // update existing map with new index
                existing.put(newHeaders.get(i), lastCol + i);
            }

            // Build a map of riceId -> sheet row indices for quick lookup (allow multiple rows per riceId)
            Map<String, List<Integer>> sheetRiceRows = new HashMap<>();
            int lastRowNum = sheet.getLastRowNum();
            for (int r = 1; r <= lastRowNum; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell first = row.getCell(0);
                String v = ExcelUtils.getStringCellValue(first);
                if (v != null && !v.isBlank()) {
                    sheetRiceRows.computeIfAbsent(v, k -> new ArrayList<>()).add(r);
                }
            }

            // Diagnostic logging
            System.out.println("DEBUG: CSV records count=" + records.size());
             // first distinct riceIds from records
             Set<String> sampleRice = records.stream().map(CsvRecord::getRiceId).filter(Objects::nonNull).limit(10).collect(Collectors.toSet());
             System.out.println("DEBUG: sample riceIds from CSV (up to 10): " + sampleRice);
             System.out.println("DEBUG: terminalsByRiceId count=" + terminalsByRiceId.size());
             System.out.println("DEBUG: template sheet riceIds count=" + sheetRiceRows.size());
             System.out.println("DEBUG: sample template riceIds (up to 10): " + sheetRiceRows.keySet().stream().limit(10).toList());

             // Fill terminal columns for EVERY row in template
             // Iterate through ALL template rows (not CSV riceIds)
             int filledTerminalCells = 0;
             int templateRowCount = sheet.getLastRowNum();

             for (int r = 1; r <= templateRowCount; r++) {
                Row templateRow = sheet.getRow(r);
                if (templateRow == null) continue;

                Cell firstCell = templateRow.getCell(0);
                String templateRiceId = ExcelUtils.getStringCellValue(firstCell);

                if (templateRiceId != null && !templateRiceId.trim().isEmpty()) {
                    // Look for terminals for this template riceId in the CSV data
                    List<String> terminals = terminalsByRiceId.get(templateRiceId.trim());

                    if (terminals == null) {
                        // Try case-insensitive lookup
                        for (Map.Entry<String, List<String>> entry : terminalsByRiceId.entrySet()) {
                            if (entry.getKey() != null && entry.getKey().trim().equalsIgnoreCase(templateRiceId.trim())) {
                                terminals = entry.getValue();
                                break;
                            }
                        }
                    }

                    if (terminals != null && !terminals.isEmpty()) {
                        // Add terminals to this template row
                        for (int i = 0; i < terminals.size(); i++) {
                            String header = "terminal_" + (i + 1);
                            Integer colIndex = existing.get(header);
                            if (colIndex == null) continue; // shouldn't happen

                            Cell cell = ExcelUtils.createCellIfAbsent(templateRow, colIndex);
                            String value = terminals.get(i);
                            if (value != null) {
                                cell.setCellValue(value);
                                filledTerminalCells++;
                            }
                        }
                    } else {
                        // No terminals found for this template riceId - leave terminal columns empty (as required)
                        // This is expected and correct behavior
                    }
                }
             }

             System.out.println("DEBUG: filled terminal cells count = " + filledTerminalCells);

            // DO NOT remove duplicate rows - keep all rows from template as required
            // The requirement is to maintain the same number of rows (2360) as the template

            System.out.println("DEBUG: template sheet riceIds count=" + sheetRiceRows.size());
            List<String> sampleKeys = new ArrayList<String>();
            int count = 0;
            for (String key : sheetRiceRows.keySet()) {
                if (count >= 10) break;
                sampleKeys.add(key);
                count++;
            }
            System.out.println("DEBUG: sample template riceIds (up to 10): " + sampleKeys);
            System.out.println("DEBUG: Total rows in template after processing: " + (sheet.getLastRowNum() + 1));


            // write workbook to a temp file then move atomically
            Path out = Path.of(outputPath);

            // Remove unwanted columns from sheet (chiave, tid, timestamp_TO_DELETE, valore, timestamp)
            List<String> colsToRemove = Arrays.asList("chiave", "tid", "timestamp_TO_DELETE", "valore", "timestamp", "rice_id", "terminal_id");
            // compute set of column indices to remove based on current header map
            Set<Integer> removeIdx = new HashSet<>();
            for (String cName : colsToRemove) {
                Integer idx = existing.get(cName);
                if (idx != null) removeIdx.add(idx);
            }
            if (!removeIdx.isEmpty()) {
                // build list of kept column indices in ascending order
                int maxCol = existing.values().stream().mapToInt(Integer::intValue).max().orElse(-1);
                List<Integer> keepIndices = new ArrayList<>();
                for (int ci = 0; ci <= maxCol; ci++) {
                    if (!removeIdx.contains(ci)) keepIndices.add(ci);
                }

                // create a temporary sheet and copy only kept columns
                String originalName = sheet.getSheetName();
                Sheet newSheet = workbook.createSheet(originalName + "_copy");
                int lastRow = sheet.getLastRowNum();
                for (int r = 0; r <= lastRow; r++) {
                    Row srcRow = sheet.getRow(r);
                    Row dstRow = newSheet.createRow(r);
                    if (srcRow == null) continue;
                    for (int j = 0; j < keepIndices.size(); j++) {
                        int srcCol = keepIndices.get(j);
                        Cell srcCell = srcRow.getCell(srcCol);
                        if (srcCell == null) continue;
                        Cell dstCell = dstRow.createCell(j);
                        copyCellValue(srcCell, dstCell);
                    }
                }
                // remove the old sheet and rename the new one to original name
                int origIndex = workbook.getSheetIndex(sheet);
                workbook.removeSheetAt(origIndex);
                int newIndex = workbook.getSheetIndex(newSheet);
                workbook.setSheetName(newIndex, originalName);

                // Safety: after creating the new sheet, ensure terminal_* values are present by rewriting them
                Sheet finalSheet = workbook.getSheetAt(newIndex);
                Map<String,Integer> finalHeaders = ExcelUtils.headerMap(finalSheet);
                // build map riceId->row index for finalSheet
                Map<String,List<Integer>> finalSheetRiceRows = new HashMap<>();
                int lastR = finalSheet.getLastRowNum();
                for (int rr = 1; rr <= lastR; rr++) {
                    Row row = finalSheet.getRow(rr);
                    if (row == null) continue;
                    String v0 = ExcelUtils.getStringCellValue(row.getCell(0));
                    if (v0 != null && !v0.isBlank()) finalSheetRiceRows.computeIfAbsent(v0, k -> new ArrayList<>()).add(rr);
                }
                // rewrite terminal values from previously computed terminalsByRiceId
                for (Map.Entry<String, List<String>> te : terminalsByRiceId.entrySet()) {
                    String riceId = te.getKey();
                    List<Integer> crow = finalSheetRiceRows.get(riceId);
                    if (crow == null || crow.isEmpty()) continue;
                    int rowIdx = crow.get(0);
                    Row dstRow = finalSheet.getRow(rowIdx);
                    if (dstRow == null) dstRow = finalSheet.createRow(rowIdx);
                    List<String> terms = te.getValue();
                    for (int i = 0; i < terms.size(); i++) {
                        String header = "terminal_" + (i + 1);
                        Integer colIdx = finalHeaders.get(header);
                        if (colIdx == null) continue;
                        String val = terms.get(i);
                        if (val == null) continue;
                        Cell cell = ExcelUtils.createCellIfAbsent(dstRow, colIdx);
                        cell.setCellValue(val);
                    }
                }

                // Ensure rice_id and terminal_id are removed: if present in finalHeaders, make a filtered copy excluding them
                List<String> mustRemove = Arrays.asList("rice_id", "terminal_id");
                boolean needSecondFilter = false;
                for (String rm : mustRemove) if (finalHeaders.containsKey(rm)) { needSecondFilter = true; break; }
                if (needSecondFilter) {
                    // compute keep columns based on finalHeaders order
                    int maxCol2 = finalHeaders.values().stream().mapToInt(Integer::intValue).max().orElse(-1);
                    List<Integer> keepIdx2 = new ArrayList<>();
                    for (int ci = 0; ci <= maxCol2; ci++) {
                        boolean remove = false;
                        for (Map.Entry<String,Integer> fh : finalHeaders.entrySet()) {
                            if (fh.getValue() == ci && mustRemove.contains(fh.getKey())) { remove = true; break; }
                        }
                        if (!remove) keepIdx2.add(ci);
                    }
                    Sheet filtered = workbook.createSheet(originalName + "_filtered");
                    int lastRow2 = finalSheet.getLastRowNum();
                    for (int r = 0; r <= lastRow2; r++) {
                        Row src = finalSheet.getRow(r);
                        Row dst = filtered.createRow(r);
                        if (src == null) continue;
                        for (int j = 0; j < keepIdx2.size(); j++) {
                            int sc = keepIdx2.get(j);
                            Cell scell = src.getCell(sc);
                            if (scell == null) continue;
                            Cell dcell = dst.createCell(j);
                            copyCellValue(scell, dcell);
                        }
                    }
                    // remove finalSheet and rename filtered to originalName
                    int fi = workbook.getSheetIndex(finalSheet);
                    workbook.removeSheetAt(fi);
                    int idx2 = workbook.getSheetIndex(filtered);
                    workbook.setSheetName(idx2, originalName);
                }
             }

            // Safety: avoid overwriting the input template file when templatePath is a filesystem path and equals outputPath
            if (!templatePath.startsWith("classpath:")) {
                try {
                    Path templateFile = Path.of(templatePath).toAbsolutePath().normalize();
                    Path outFile = out.toAbsolutePath().normalize();
                    if (Files.exists(templateFile) && templateFile.equals(outFile)) {
                        // compute alternative output file name with timestamp
                        String name = outFile.getFileName().toString();
                        String base = name;
                        String ext = "";
                        int idx = name.lastIndexOf('.');
                        if (idx > 0) {
                            base = name.substring(0, idx);
                            ext = name.substring(idx);
                        }
                        String newName = base + "-generated-" + Instant.now().toEpochMilli() + ext;
                        out = outFile.getParent().resolve(newName);
                        System.out.println("Output path equals template path; using alternate output: " + out);
                     }
                 } catch (Exception ex) {
                     // ignore and proceed with provided outputPath
                 }
             }

            Path parent = out.toAbsolutePath().getParent();
            if (parent == null) {
                // fallback to system temp directory if no parent is available
                parent = Path.of(System.getProperty("java.io.tmpdir"));
            }
            // Ensure destination directory exists before creating temp file in it
            Files.createDirectories(parent);
            Path tmp = Files.createTempFile(parent, "estrazione-", ".tmp.xlsx");
            try (OutputStream fos = Files.newOutputStream(tmp)) {
                workbook.write(fos);
            }
            Files.move(tmp, out, java.nio.file.StandardCopyOption.REPLACE_EXISTING, java.nio.file.StandardCopyOption.ATOMIC_MOVE);
        }
    }

    private InputStream openResource(String path) throws IOException {
        if (path == null) return null;
        if (path.startsWith("classpath:")) {
            String cp = path.substring("classpath:".length());
            Resource r = new ClassPathResource(cp);
            if (!r.exists()) return null;
            return r.getInputStream();
        }
        Path p = Path.of(path);
        if (!Files.exists(p)) return null;
        return Files.newInputStream(p);
    }

    private List<String[]> readCsv(String csvPath) throws IOException {
        InputStream in = openResource(csvPath);
        if (in == null) throw new FileNotFoundException("CSV not found: " + csvPath);
        try (Reader reader = new InputStreamReader(in)) {
            try (CSVReader csvReader = new CSVReaderBuilder(reader)
                    .withCSVParser(new CSVParserBuilder().withSeparator(';').build())
                    .build()) {
                try {
                    return csvReader.readAll();
                } catch (CsvException e) {
                    throw new IOException("CSV parsing error", e);
                }
            }
        }
    }

    private void copyCellValue(Cell src, Cell dst) {
        dst.setCellType(src.getCellType());
        switch (src.getCellType()) {
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case NUMERIC:
                dst.setCellValue(src.getNumericCellValue());
                break;
            case FORMULA:
                dst.setCellFormula(src.getCellFormula());
                break;
            case BLANK:
                dst.setBlank();
                break;
            case ERROR:
                dst.setCellErrorValue(src.getErrorCellValue());
                break;
            default:
                throw new IllegalArgumentException("Unsupported cell type: " + src.getCellType());
        }
    }
}
