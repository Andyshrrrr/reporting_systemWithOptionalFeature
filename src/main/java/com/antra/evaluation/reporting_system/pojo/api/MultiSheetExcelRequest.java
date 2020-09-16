package com.antra.evaluation.reporting_system.pojo.api;

import org.springframework.lang.NonNull;

import java.util.List;

public class MultiSheetExcelRequest extends ExcelRequest{
    @NonNull
    private List<String> headers;
    @NonNull
    private List<List<Object>> data;
    @NonNull
    private String splitBy;
    private String description;

    public List<List<Object>> getData() {
        return data;
    }

    public void setData(List<List<Object>> data) {
        this.data = data;
    }

    public String getSplitBy() {
        return splitBy;
    }

    public void setSplitBy(String splitBy) {
        this.splitBy = splitBy;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public List<String> getHeaders() {
        return headers;
    }

    public void setHeaders(List<String> headers) {
        this.headers = headers;
    }
}
