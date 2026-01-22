package com.example.estrazione_win10_sales.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

@Component
@ConfigurationProperties(prefix = "excel")
public class ExcelProperties {

    /** Path to template, supports classpath: prefix */
    private String templatePath = "classpath:inputfile/template.xlsx";

    /** Output file path */
    private String outputPath = "target/estrazione_output.xlsx";

    /** CSV input path */
    private String csvPath = "classpath:inputfile/estrazione.csv";

    public String getTemplatePath() {
        return templatePath;
    }

    public void setTemplatePath(String templatePath) {
        this.templatePath = templatePath;
    }

    public String getOutputPath() {
        return outputPath;
    }

    public void setOutputPath(String outputPath) {
        this.outputPath = outputPath;
    }

    public String getCsvPath() {
        return csvPath;
    }

    public void setCsvPath(String csvPath) {
        this.csvPath = csvPath;
    }
}
