package com.example.estrazione_win10_sales;

import com.example.estrazione_win10_sales.config.ExcelProperties;
import com.example.estrazione_win10_sales.service.CsvToExcelService;
import org.springframework.boot.CommandLineRunner;

public class CsvToExcelRunner implements CommandLineRunner {

    private final CsvToExcelService service;
    private final ExcelProperties props;

    public CsvToExcelRunner(CsvToExcelService service, ExcelProperties props) {
        this.service = service;
        this.props = props;
    }

    @Override
    public void run(String... args) throws Exception {
        // Run conversion on startup; if you want to disable, add a property later
        try {
            service.run(props.getCsvPath(), props.getTemplatePath(), props.getOutputPath());
            System.out.println("CsvToExcelRunner: conversion completed, output=" + props.getOutputPath());
        } catch (Exception e) {
            System.err.println("CsvToExcelRunner: conversion failed: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
