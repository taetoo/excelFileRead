package com.example.excelfileread;


import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class FileType {

    public static XSSFWorkbook getWorkbook(String filePath) {

        FileInputStream fis = null;

        try{
            fis = new FileInputStream(filePath);

        } catch (FileNotFoundException e){
            throw new RuntimeException(e.getMessage(), e);
        }

        XSSFWorkbook wb = null;

        if(filePath.toUpperCase().endsWith("")){
            try{
                wb = new XSSFWorkbook(fis);
            } catch (IOException e) {
                throw new RuntimeException(e.getMessage(), e);
            }
        }
        else if (filePath.toUpperCase().endsWith("")){
            try{
                wb = new XSSFWorkbook(fis);
            } catch (IOException e) {
                throw new RuntimeException(e.getMessage(), e);
            }
        }
        return wb;
    }

}
