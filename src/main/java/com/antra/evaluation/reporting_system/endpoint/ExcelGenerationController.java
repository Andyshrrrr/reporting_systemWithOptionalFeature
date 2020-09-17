package com.antra.evaluation.reporting_system.endpoint;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.ExcelResponse;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import com.antra.evaluation.reporting_system.service.ExcelService;
import io.swagger.annotations.ApiOperation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.util.FileCopyUtils;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.*;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@RestController
public class ExcelGenerationController {

    private static final Logger log = LoggerFactory.getLogger(ExcelGenerationController.class);
    ExcelService excelService;

    @Autowired
    public ExcelGenerationController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @PostMapping("/excel")
    @ApiOperation("Generate Excel")
    public ResponseEntity<ExcelResponse> createExcel(@RequestBody @Validated ExcelRequest request) throws IOException {
        log.info("createExcel() start");
        ExcelResponse response = new ExcelResponse();
        ExcelFile file = excelService.CreateOneSheetExcel(request);
        if(file ==  null) {
            log.error("createExcel() error: fail to save into repository");
            throw new IOException("Fail to save into repository");
        }
        response.setFileId(file.getID());
        response.setGeneratedTime(file.getGenerateTime());
        response.setFileSize(file.getFileSize());
        response.setDownloadLink(file.getPathOfFile());
        log.info("createExcel() end");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    @PostMapping("/excel/auto")
    @ApiOperation("Generate Multi-Sheet Excel Using Split field")
    public ResponseEntity<ExcelResponse> createMultiSheetExcel(@RequestBody @Validated MultiSheetExcelRequest request) throws IOException{
        log.info("createMultiSheetExcel() start");
        ExcelResponse response = new ExcelResponse();
        ExcelFile file = excelService.CreatMultiSheetExcel(request);
        if(file ==  null) {
            log.error("createMultiSheet() error: Fail to save into repository");
            throw new RuntimeException("Fail to save into repository");
        }
        response.setFileId(file.getID());
        response.setGeneratedTime(file.getGenerateTime());
        response.setFileSize(file.getFileSize());
        response.setDownloadLink(file.getPathOfFile());
        log.info("createMultiSheetExcel() end");
        return new ResponseEntity<>(response, HttpStatus.OK);
    }

    @GetMapping("/excel")
    @ApiOperation("List all existing files")
    public ResponseEntity<List<ExcelResponse>> listExcels() {
        log.info("listExcels() start");
        var responses = new ArrayList<ExcelResponse>();
        List<ExcelFile> files = excelService.getAllInRepository();
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
        ExcelFile deleteFile = excelService.deleteByID(id);
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
    public void exceptionThrown2(MethodArgumentNotValidException e) {
        e.printStackTrace();
    }

    @ExceptionHandler
    @ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
    public void exceptionThrown3(Exception e) {
        e.printStackTrace();
    }
}