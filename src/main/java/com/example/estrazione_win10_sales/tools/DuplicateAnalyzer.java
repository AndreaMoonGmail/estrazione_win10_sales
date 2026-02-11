package com.example.estrazione_win10_sales.tools;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.util.*;

public class DuplicateAnalyzer {

    public static void main(String[] args) throws Exception {
        System.out.println("=== ANALISI DUPLICATI NEL TEMPLATE ===");

        // Leggi il template
        Path templatePath = Paths.get("src/main/resources/inputfile/template.xlsx");

        try (InputStream in = new FileInputStream(templatePath.toFile());
             Workbook wb = new XSSFWorkbook(in)) {

            Sheet sheet = wb.getSheetAt(0);

            // Map per tracciare idrice -> lista di righe dove compare
            Map<String, List<Integer>> idriceToRows = new LinkedHashMap<>();

            // Scorri tutte le righe (escludi header)
            int totalRows = sheet.getLastRowNum();
            System.out.println("Numero totale righe nel template: " + (totalRows + 1)); // +1 perché include header
            System.out.println("Righe dati (escluso header): " + totalRows);

            for (int r = 1; r <= totalRows; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Cell firstCell = row.getCell(0); // Colonna A = idrice
                String idrice = getStringCellValue(firstCell);

                if (idrice != null && !idrice.trim().isEmpty()) {
                    if (!idriceToRows.containsKey(idrice)) {
                        idriceToRows.put(idrice, new ArrayList<>());
                    }
                    idriceToRows.get(idrice).add(r);
                }
            }

            System.out.println("Idrice univoci trovati nel template: " + idriceToRows.size());

            // Trova i duplicati
            List<String> duplicatedIdrice = new ArrayList<>();
            int totalDuplicatedRows = 0;

            for (Map.Entry<String, List<Integer>> entry : idriceToRows.entrySet()) {
                String idrice = entry.getKey();
                List<Integer> rows = entry.getValue();

                if (rows.size() > 1) {
                    duplicatedIdrice.add(idrice);
                    totalDuplicatedRows += (rows.size() - 1); // -1 perché una riga la teniamo
                }
            }

            System.out.println("\n=== STATISTICHE DUPLICATI ===");
            System.out.println("Idrice duplicati: " + duplicatedIdrice.size());
            System.out.println("Righe duplicate totali: " + totalDuplicatedRows);
            System.out.println("Righe che verrebbero rimosse: " + totalDuplicatedRows);
            System.out.println("Righe finali se rimuovi duplicati: " + (totalRows - totalDuplicatedRows + 1)); // +1 per header

            System.out.println("\n=== PRIMI 10 IDRICE DUPLICATI ===");

            int count = 0;
            for (String idrice : duplicatedIdrice) {
                if (count >= 10) break;

                List<Integer> rows = idriceToRows.get(idrice);
                System.out.println("Idrice: '" + idrice + "' trovato nelle righe: " + rows);

                // Mostra il contenuto di alcune colonne per ogni riga duplicata
                for (int rowNum : rows) {
                    Row row = sheet.getRow(rowNum);
                    if (row != null) {
                        Integer osVersionColIndex = getColumnIndex(sheet, "osVersion");
                        String osVersion = osVersionColIndex != null ?
                            getStringCellValue(row.getCell(osVersionColIndex)) : "N/A";
                        String descrizione = getStringCellValue(row.getCell(1)); // Assuming column B is description

                        String descSub = descrizione != null ?
                            descrizione.substring(0, Math.min(50, descrizione.length())) + "..." : "N/A";

                        System.out.println("  Riga " + (rowNum + 1) + ": osVersion='" + osVersion + "', Descrizione='" + descSub + "'");
                    }
                }
                System.out.println();
                count++;
            }

            if (duplicatedIdrice.size() > 10) {
                System.out.println("... e altri " + (duplicatedIdrice.size() - 10) + " idrice duplicati.");
            }

            System.out.println("\n=== ESEMPIO DI IDRICE DUPLICATO PER TEST MANUALE ===");
            if (!duplicatedIdrice.isEmpty()) {
                String firstDuplicate = duplicatedIdrice.getFirst();
                System.out.println("Idrice da testare manualmente: " + firstDuplicate);
                System.out.println("Righe nel template: " + idriceToRows.get(firstDuplicate));
                System.out.println("Puoi cercare questo idrice nel file template.xlsx per verificare manualmente.");
            }
        }
    }

    private static Integer getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        if (headerRow == null) return null;

        for (Cell cell : headerRow) {
            String headerValue = getStringCellValue(cell);
            if (columnName.equals(headerValue)) {
                return cell.getColumnIndex();
            }
        }
        return null;
    }

    /**
     * Utility method to get string value from Excel cell (replaces ExcelUtils.getStringCellValue)
     */
    private static String getStringCellValue(Cell c) {
        if (c == null) return null;
        if (c.getCellType() == CellType.STRING) return c.getStringCellValue();
        if (c.getCellType() == CellType.NUMERIC) {
            // Format numbers properly to avoid scientific notation
            double numericValue = c.getNumericCellValue();
            if (numericValue == Math.floor(numericValue)) {
                // It's an integer
                return String.valueOf((long) numericValue);
            } else {
                // It's a decimal
                DecimalFormat df = new DecimalFormat("#.##########");
                return df.format(numericValue);
            }
        }
        if (c.getCellType() == CellType.BOOLEAN) return String.valueOf(c.getBooleanCellValue());
        if (c.getCellType() == CellType.FORMULA) return c.getCellFormula();
        return null;
    }
}
