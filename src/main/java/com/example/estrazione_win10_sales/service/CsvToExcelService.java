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
                    .toList());

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

             // Fill terminal columns for each riceId found in the sheet
             int filledTerminalCells = 0;
             for (Map.Entry<String, List<String>> e : terminalsByRiceId.entrySet()) {
                String riceId = e.getKey();
                List<String> terms = e.getValue();
                List<Integer> candidateRows = sheetRiceRows.get(riceId);
                Integer rowIndex = null;
                if (candidateRows != null && !candidateRows.isEmpty()) {
                    // choose keeper row: prefer the row that has a non-empty osVersion cell
                    Integer osVerCol = existing.get("osVersion");
                    if (osVerCol != null) {
                        for (Integer rr : candidateRows) {
                            Row row = sheet.getRow(rr);
                            if (row == null) continue;
                            Cell c = row.getCell(osVerCol);
                            String sv = ExcelUtils.getStringCellValue(c);
                            if (sv != null && !sv.isBlank()) {
                                rowIndex = rr;
                                break;
                            }
                        }
                    }
                    // fallback: use first candidate
                    if (rowIndex == null) rowIndex = candidateRows.get(0);
                }
                // best-effort: if still null, try robust matching via keys present in sheetRiceRows map
                if (rowIndex == null) {
                    // try trimmed case-insensitive key lookup
                    for (Map.Entry<String, List<Integer>> me : sheetRiceRows.entrySet()) {
                        String key = me.getKey();
                        if (key != null && riceId != null && key.trim().equalsIgnoreCase(riceId.trim())) {
                            List<Integer> list = me.getValue();
                            if (!list.isEmpty()) rowIndex = list.get(0);
                            break;
                        }
                    }
                }
                if (rowIndex == null) {
                    // riceId not present in template
                    continue;
                }
                Row excelRow = sheet.getRow(rowIndex);
                if (excelRow == null) excelRow = sheet.createRow(rowIndex);
                // write each terminal value into the corresponding terminal_x column
                for (int i = 0; i < terms.size(); i++) {
                    String header = "terminal_" + (i + 1);
                    Integer colIndex = existing.get(header);
                    if (colIndex == null) continue; // shouldn't happen
                    Cell cell = ExcelUtils.createCellIfAbsent(excelRow, colIndex);
                    String value = terms.get(i);
                    if (value != null) cell.setCellValue(value);
                    if (value != null) filledTerminalCells++;
                }
             }

             System.out.println("DEBUG: filled terminal cells count = " + filledTerminalCells);

            // Now remove duplicate rows per riceId, keeping only the keeper selected above (osVersion row)
            // Collect all rows to remove
            List<Integer> rowsToRemove = new ArrayList<>();
            for (Map.Entry<String, List<Integer>> me : sheetRiceRows.entrySet()) {
                List<Integer> list = me.getValue();
                if (list.size() <= 1) continue;
                // determine keeper as above
                Integer keeper = null;
                Integer osVerCol = existing.get("osVersion");
                if (osVerCol != null) {
                    for (Integer rr : list) {
                        Row row = sheet.getRow(rr);
                        if (row == null) continue;
                        String sv = ExcelUtils.getStringCellValue(row.getCell(osVerCol));
                        if (sv != null && !sv.isBlank()) { keeper = rr; break; }
                    }
                }
                if (keeper == null) keeper = list.get(0);
                for (Integer rr : list) {
                    if (!rr.equals(keeper)) rowsToRemove.add(rr);
                }
            }
            if (!rowsToRemove.isEmpty()) {
                // sort descending to safely remove rows and shift
                rowsToRemove.sort(Comparator.reverseOrder());
                for (Integer rr : rowsToRemove) {
                    int last = sheet.getLastRowNum();
                    Row row = sheet.getRow(rr);
                    if (row != null) {
                        sheet.removeRow(row);
                        if (rr < last) {
                            sheet.shiftRows(rr + 1, last, -1);
                        }
                    }
                }
            }

            // After shifting rows, rebuild the map of riceId -> sheet row indices because indices changed
            sheetRiceRows.clear();
            lastRowNum = sheet.getLastRowNum();
            for (int r = 1; r <= lastRowNum; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell first = row.getCell(0);
                String v = ExcelUtils.getStringCellValue(first);
                if (v != null && !v.isBlank()) {
                    sheetRiceRows.computeIfAbsent(v, k -> new ArrayList<>()).add(r);
                }
            }

            System.out.println("DEBUG: rebuilt template sheet riceIds count=" + sheetRiceRows.size());
            System.out.println("DEBUG: sample rebuilt template riceIds (up to 10): " + sheetRiceRows.keySet().stream().limit(10).toList());

            // Also keep existing CSV-to-sheet behavior for other CSV headers (map csv rows to excel rows by index)
            // write data rows: map each CSV row to the corresponding template row by riceId
            for (int i = 1; i < rows.size(); i++) {
                String[] data = rows.get(i);
                if (data == null || data.length == 0) continue;
                // sanitize riceId (trim and remove surrounding quotes)
                String riceIdRaw = data[0];
                if (riceIdRaw == null) continue;
                String riceId = riceIdRaw.trim();
                if (riceId.startsWith("\"") && riceId.endsWith("\"") && riceId.length() >= 2) {
                    riceId = riceId.substring(1, riceId.length() - 1).trim();
                }
                if (riceId.isBlank()) continue;

                // find corresponding template row index for this riceId
                Integer rowIndex = null;
                List<Integer> candidateRows = sheetRiceRows.get(riceId);
                if (candidateRows != null && !candidateRows.isEmpty()) {
                    rowIndex = candidateRows.get(0);
                } else {
                    // fallback: trimmed case-insensitive lookup
                    for (Map.Entry<String, List<Integer>> me : sheetRiceRows.entrySet()) {
                        String key = me.getKey();
                        if (key != null && key.trim().equalsIgnoreCase(riceId.trim())) {
                            List<Integer> list = me.getValue();
                            if (!list.isEmpty()) rowIndex = list.get(0);
                            break;
                        }
                    }
                }
                if (rowIndex == null) continue; // riceId not present in template

                Row r = sheet.getRow(rowIndex);
                if (r == null) r = sheet.createRow(rowIndex);

                // write values for each CSV header into the found template row
                for (int col = 0; col < headers.length; col++) {
                    String header = headers[col];
                    if (header == null || header.isBlank()) continue;
                    Integer colIndex = existing.get(header);
                    if (colIndex == null) continue; // shouldn't happen
                    Cell cell = ExcelUtils.createCellIfAbsent(r, colIndex);
                    String value = data.length > col ? data[col] : null;
                    if (value != null) {
                        // strip surrounding quotes
                        if (value.startsWith("\"") && value.endsWith("\"")) {
                            value = value.substring(1, value.length()-1);
                        }
                        cell.setCellValue(value);
                    }
                }
            }

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
