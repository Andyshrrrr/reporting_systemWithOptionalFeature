package com.antra.evaluation.reporting_system.repo;

import com.antra.evaluation.reporting_system.pojo.report.ExcelFile;
import org.springframework.stereotype.Repository;

import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

@Repository
public class ExcelRepositoryImpl implements ExcelRepository {
    Map<String, ExcelFile> excelData = new ConcurrentHashMap<>();

    @Override
    public Optional<ExcelFile> getFileById(String id) {
        return Optional.ofNullable(excelData.get(id));
    }

    @Override
    public ExcelFile saveFile(ExcelFile file) {
        try {
            String ID = file.getID();
            excelData.put(ID, file);
        }
        catch (Exception e) {
            e.printStackTrace();
            return null;
        }
        return file;
    }

    @Override
    public ExcelFile deleteFile(String id) {
        ExcelFile file = excelData.get(id);
        if(file == null) {
            return null;
        }
        excelData.remove(id);
        file.getFile().delete();
        return file;
    }

    @Override
    public List<ExcelFile> getFiles() {
        List<ExcelFile> files = new ArrayList<>();
        Iterator iterator = excelData.values().iterator();
        while(iterator.hasNext()) {
            ExcelFile file = (ExcelFile)iterator.next();
            files.add(file);
        }
        return files;
    }
}

