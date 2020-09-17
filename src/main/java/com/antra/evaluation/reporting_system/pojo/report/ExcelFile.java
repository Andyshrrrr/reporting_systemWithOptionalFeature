package com.antra.evaluation.reporting_system.pojo.report;

import com.antra.evaluation.reporting_system.service.ExcelGenerationServiceImpl;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;

public class ExcelFile {
    ExcelData data;
    File file;
    String ID;
    LocalDateTime GenerateTime;
    long fileSize;

    public void setData(ExcelData data) throws IOException {
        this.data = data;
        file = new ExcelGenerationServiceImpl().generateExcelReport(data);
        ID = data.getTitle();
        GenerateTime = data.getGeneratedTime();
        fileSize = file.length();
    }

    //for mock test usage,
    public ExcelFile(ExcelData data, File file, String ID, LocalDateTime generateTime, long fileSize) {
        this.data = data;
        this.file = file;
        this.ID = ID;
        GenerateTime = generateTime;
        this.fileSize = fileSize;
    }

    public ExcelFile() {
    }

    public ExcelData getData() {
        return data;
    }

    public String getPathOfFile() {
        return file.getAbsolutePath();
    }

    public String getID() {
        return ID;
    }

    public LocalDateTime getGenerateTime() {
        return GenerateTime;
    }

    public long getFileSize() {
        return fileSize;
    }

    public File getFile() {
        return file;
    }
}
