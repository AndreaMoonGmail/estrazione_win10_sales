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
        } catch (Exception e) {
            System.err.println("Estrazione: conversion failed: " + e.getMessage());
            e.printStackTrace();
            // rilanciare per far fallire l'app se necessario
            throw e;
        }
    }
}
