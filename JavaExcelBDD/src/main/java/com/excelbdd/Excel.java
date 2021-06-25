package com.excelbdd;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
    public final FileInputStream excelFileInputStream;

    public Excel(FileInputStream excelFileInputStream) {
        this.excelFileInputStream = excelFileInputStream;
    }

    public static Excel open(String excelPath) throws FileNotFoundException {
        return new Excel(new FileInputStream(excelPath));
    }

    public Sheet openSheet(String worksheetName) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
        XSSFSheet sheet = workbook.getSheet(worksheetName);
        if (sheet == null) {
            workbook.close();
            excelFileInputStream.close();
            throw new IOException(worksheetName + " does not exist.");
        }
        return new Sheet(sheet);
    }

    public void close() throws IOException {
        excelFileInputStream.close();
    }
}
