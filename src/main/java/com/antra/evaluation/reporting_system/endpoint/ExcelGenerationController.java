package com.antra.evaluation.reporting_system.endpoint;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.ExcelResponse;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataHeader;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataSheet;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import com.antra.evaluation.reporting_system.repo.ExcelRepositoryImpl;
import com.antra.evaluation.reporting_system.service.ExcelService;
import io.swagger.annotations.ApiOperation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.util.FileCopyUtils;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.stream.Collectors;

@RestController
public class ExcelGenerationController {

    private static final Logger log = LoggerFactory.getLogger(ExcelGenerationController.class);
    ExcelService excelService;
    ExcelRepositoryImpl repository;
    int dataCounter = 1;

    @Autowired
    public ExcelGenerationController(ExcelService excelService, ExcelRepositoryImpl repository) {
        this.excelService = excelService;
        this.repository = repository;
    }

    @PostMapping("/excel")
    @ApiOperation("Generate Excel")
    public ResponseEntity<ExcelResponse> createExcel(@RequestBody @Validated ExcelRequest request) throws IOException {
        log.info("createExcel() start");
        ExcelResponse response = new ExcelResponse();
        ExcelData data = new ExcelData();
        data.setGeneratedTime(LocalDateTime.now());
        ExcelDataSheet sheet1 = new ExcelDataSheet();
        List<ExcelDataSheet> sheets = new ArrayList<>();
        var AllHeader = new ArrayList<ExcelDataHeader>();
        for(String s:request.getHeaders()){
            ExcelDataHeader header = new ExcelDataHeader();
            header.setName(s);
            AllHeader.add(header);
        }
        sheet1.setHeaders(AllHeader);
        sheet1.setDataRows(request.getData());
        sheet1.setTitle("sheet1");
        sheets.add(sheet1);
        data.setSheets(sheets);
        data.setTitle("data"+dataCounter);
        dataCounter++;
        ExcelFile file = new ExcelFile();
        file.setData(data);
        response.setFileId(file.getID());
        response.setGeneratedTime(file.getGenerateTime());
        response.setFileSize(file.getFileSize());
        response.setDownloadLink(file.getPathOfFile());
        if(repository.saveFile(file) ==  null) {
            log.error("createExcel() error: fail to save into repository");
            throw new RuntimeException("Fail to save into repository");
        }
        log.info("createExcel() end");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    @PostMapping("/excel/auto")
    @ApiOperation("Generate Multi-Sheet Excel Using Split field")
    public ResponseEntity<ExcelResponse> createMultiSheetExcel(@RequestBody @Validated MultiSheetExcelRequest request) throws IOException{
        log.info("createMultiSheetExcel() start");
        ExcelResponse response = new ExcelResponse();
        ExcelData data = new ExcelData();
        data.setGeneratedTime(LocalDateTime.now());
        List<ExcelDataSheet> sheets = new ArrayList<>();
        var AllHeaders = new ArrayList<ExcelDataHeader>();
        int SplitPos = request.getHeaders().indexOf(request.getSplitBy());
        for(String s:request.getHeaders()){
            ExcelDataHeader header = new ExcelDataHeader();
            header.setName(s);
            AllHeaders.add(header);
        }
        HashSet<String> allSheetsName = new HashSet<String>();
        for(List<Object> row : request.getData()) {
            allSheetsName.add((String)row.get(SplitPos));
        }
        List<String> allSheetsNameDes = allSheetsName.stream().sorted().collect(Collectors.toList());
        for(String sheetName: allSheetsNameDes) {
            ExcelDataSheet sheet = new ExcelDataSheet();
            sheet.setHeaders(AllHeaders);
            sheet.setTitle(sheetName);
            List<List<Object>> rows = new ArrayList<>();
            for(List<Object> row : request.getData()) {
                if (row.get(SplitPos).equals(sheetName)) {
                    rows.add(row);
                }
            }
            sheet.setDataRows(rows);
            sheets.add(sheet);
        }
        data.setTitle("data" + dataCounter);
        data.setSheets(sheets);
        dataCounter++;
        ExcelFile file = new ExcelFile();
        file.setData(data);
        response.setFileId(file.getID());
        response.setGeneratedTime(file.getGenerateTime());
        response.setFileSize(file.getFileSize());
        response.setDownloadLink(file.getPathOfFile());
        if(repository.saveFile(file) ==  null) {
            log.error("createMultiSheet() error: Fail to save into repository");
            throw new RuntimeException("Fail to save into repository");
        }
        log.info("createMultiSheetExcel() end");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    @GetMapping("/excel")
    @ApiOperation("List all existing files")
    public ResponseEntity<List<ExcelResponse>> listExcels() {
        log.info("listExcels() start");
        var responses = new ArrayList<ExcelResponse>();
        List<ExcelFile> files = repository.getFiles();
        if(files.size() < 1) {
            log.warn("listExcels() warning: no files saved");
        }
        for(ExcelFile file:files) {
            ExcelResponse response = new ExcelResponse();
            response.setDownloadLink(file.getPathOfFile());
            response.setFileId(file.getID());
            response.setGeneratedTime(file.getGenerateTime());
            response.setFileSize(file.getFileSize());
            responses.add(response);
        }
        log.info("listExcels() end");
        return new ResponseEntity<>(responses, HttpStatus.OK);
    }

    @GetMapping("/excel/{id}/content")
    public void downloadExcel(@PathVariable String id, HttpServletResponse response) throws IOException {
        log.info("downloadExcel start");
        InputStream fis = excelService.getExcelBodyById(id);
        //Optional<ExcelFile> file = repository.getFileById(id);
        String nameOfFile = id+".xlsx";
        String responseIn = "attachment; filename=\""+nameOfFile+"\"";
        response.setHeader("Content-Type","application/vnd.ms-excel");
        response.setHeader("Content-Disposition","attachment; filename=\"name_of_excel_file.xls\"");
        response.setHeader("Content-Disposition",responseIn);
        FileCopyUtils.copy(fis, response.getOutputStream());
        log.info("downloadExcel end");
    }

    @DeleteMapping("/excel/{id}")
    public ResponseEntity<ExcelResponse> deleteExcel(@PathVariable String id) {
        log.info("deleteExcel start");
        var response = new ExcelResponse();
        ExcelFile deleteFile = repository.deleteFile(id);
        if(deleteFile == null) {
            log.warn("deleteExcel: the String id is not exist");
            return new ResponseEntity<>(response, HttpStatus.BAD_REQUEST);
        }
        response.setFileSize(deleteFile.getFileSize());
        response.setGeneratedTime(deleteFile.getGenerateTime());
        response.setFileId(id);
        response.setDownloadLink(deleteFile.getPathOfFile());
        log.info("deleteExcel end");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    @ExceptionHandler
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public void exceptionThrown(IOException e){
        e.printStackTrace();
    }
    @ExceptionHandler
    @ResponseStatus(HttpStatus.BAD_REQUEST)
    public void exceptionHandle(Exception e) {
        e.printStackTrace();
    }
}