package com.antra.evaluation.reporting_system.pojo.report;

import java.time.LocalDateTime;
import java.util.List;

public class ExcelData {
    private String title;
    private LocalDateTime generatedTime;
    private List<ExcelDataSheet> sheets;

    //for mock test usage
    public ExcelData(String title, LocalDateTime generatedTime, List<ExcelDataSheet> sheets) {
        this.title = title;
        this.generatedTime = generatedTime;
        this.sheets = sheets;
    }

    public ExcelData() {
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public LocalDateTime getGeneratedTime() {
        return generatedTime;
    }

    public void setGeneratedTime(LocalDateTime generatedTime) {
        this.generatedTime = generatedTime;
    }

    public List<ExcelDataSheet> getSheets() {
        return sheets;
    }

    public void setSheets(List<ExcelDataSheet> sheets) {
        this.sheets = sheets;
    }
}
