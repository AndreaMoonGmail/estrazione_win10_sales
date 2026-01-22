package com.example.estrazione_win10_sales;

import com.example.estrazione_win10_sales.config.ExcelProperties;
import com.example.estrazione_win10_sales.service.CsvToExcelService;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class EstrazioneWin10SalesApplication implements CommandLineRunner {

    private final CsvToExcelService service;
    private final ExcelProperties props;

    public EstrazioneWin10SalesApplication(CsvToExcelService service, ExcelProperties props) {
        this.service = service;
        this.props = props;
    }

    public static void main(String[] args) {
        SpringApplication.run(EstrazioneWin10SalesApplication.class, args);
    }

    @Override
    public void run(String... args) throws Exception {
        // avviare il processo principale da qui
        try {
            service.run(props.getCsvPath(), props.getTemplatePath(), props.getOutputPath());
            System.out.println("Estrazione: conversion completed, output=" + props.getOutputPath());

            // Diagnostic: inspect the produced Excel and write a diagnostics file
            try {
                java.nio.file.Path out = java.nio.file.Path.of(props.getOutputPath());
                if (java.nio.file.Files.exists(out)) {
                    try (java.io.InputStream in = java.nio.file.Files.newInputStream(out)) {
                        org.apache.poi.ss.usermodel.Workbook wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(in);
                        org.apache.poi.ss.usermodel.Sheet s = wb.getSheetAt(0);
                        java.util.Map<String, Integer> headers = com.example.estrazione_win10_sales.util.ExcelUtils.headerMap(s);
                        StringBuilder sb = new StringBuilder();
                        sb.append("Headers: ").append(headers).append(System.lineSeparator());
                        // print first 5 rows
                        int last = Math.min(5, s.getLastRowNum());
                        for (int r = 1; r <= last; r++) {
                            org.apache.poi.ss.usermodel.Row row = s.getRow(r);
                            if (row == null) continue;
                            for (java.util.Map.Entry<String, Integer> e : headers.entrySet()) {
                                String h = e.getKey();
                                int ci = e.getValue();
                                String val = com.example.estrazione_win10_sales.util.ExcelUtils.getStringCellValue(row.getCell(ci));
                                sb.append(h).append(" -> '").append(val).append("'\n");
                            }
                            sb.append(System.lineSeparator());
                        }
                        java.nio.file.Files.writeString(java.nio.file.Path.of("diagnostics_output.txt"), sb.toString());
                    }
                }
            } catch (Exception ex) {
                System.err.println("Diagnostic write failed: " + ex.getMessage());
            }

        } catch (Exception e) {
            System.err.println("Estrazione: conversion failed: " + e.getMessage());
            e.printStackTrace();
            // rilanciare per far fallire l'app se necessario
            throw e;
        }
    }
}
