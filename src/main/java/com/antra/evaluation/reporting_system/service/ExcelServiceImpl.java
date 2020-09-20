package com.antra.evaluation.reporting_system.service;

import com.antra.evaluation.reporting_system.pojo.api.ExcelRequest;
import com.antra.evaluation.reporting_system.pojo.api.ExcelResponse;
import com.antra.evaluation.reporting_system.pojo.api.MultiSheetExcelRequest;
import com.antra.evaluation.reporting_system.pojo.report.ExcelData;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataHeader;
import com.antra.evaluation.reporting_system.pojo.report.ExcelDataSheet;
import com.antra.evaluation.reporting_system.repo.ExcelRepository;
import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    ExcelRepository excelRepository;

    @Override
    public InputStream getExcelBodyById(String id) {

        Optional<ExcelFile> fileInfo = excelRepository.getFileById(id);
       // if (fileInfo.isPresent()) {
            String path = id + ".xlsx";
            File file = new File(path);
            try {
                return new FileInputStream(file);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            }
      //  }
        return null;
    }

    @Override
    public ExcelFile CreateOneSheetExcel(ExcelRequest request) throws IOException{
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
        data.setTitle("data"+(excelRepository.getSize()+1));
        ExcelFile file = new ExcelFile();
        file.setData(data);
        return excelRepository.saveFile(file);
    }

    @Override
    public ExcelFile CreatMultiSheetExcel(MultiSheetExcelRequest request) throws IOException {
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
        HashSet<String> allSheetsName = new HashSet<>();
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
        data.setTitle("data" + (excelRepository.getSize()+1));
        data.setSheets(sheets);
        ExcelFile file = new ExcelFile();
        file.setData(data);
        return excelRepository.saveFile(file);
    }

    @Override
    public List<ExcelFile> getAllInRepository() {
        return excelRepository.getFiles();
    }

    @Override
    public ExcelFile deleteByID(String id) {
        return excelRepository.deleteFile(id);
    }

    @Override
    public FileInputStream getBatchBodyById(List<String> id) throws IOException,FileNotFoundException{
        //System.out.println(id);
        List<String> srcFiles = id.stream().filter(w->excelRepository.getFileById(w).isPresent()==true).map(w->{return (w + ".xlsx");}).collect(Collectors.toList());
        System.out.println(srcFiles);
        FileOutputStream fos = new FileOutputStream("compresses.zip");
        ZipOutputStream zipOut = new ZipOutputStream(fos);
        for(String file:srcFiles) {
            File fileToZip = new File(file);
            FileInputStream fis = new FileInputStream(file);
            ZipEntry zipEntry = new ZipEntry(fileToZip.getName());
            zipOut.putNextEntry(zipEntry);
            byte[] bytes = new byte[1024];
            int length;
            while((length = fis.read(bytes)) >= 0) {
                zipOut.write(bytes,0,length);
            }
            fis.close();
        }
        zipOut.close();
        return new FileInputStream("compresses.zip");

    }
}
