package com.antra.evaluation.reporting_system.service;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface ExcelService {
    InputStream getExcelBodyById(String id);

    ExcelFile CreateOneSheetExcel(ExcelRequest request) throws IOException;

    ExcelFile CreatMultiSheetExcel(MultiSheetExcelRequest request) throws IOException;

    List<ExcelFile> getAllInRepository();

    ExcelFile deleteByID(String id);
}
