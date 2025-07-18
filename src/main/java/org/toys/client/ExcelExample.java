package org.toys.client;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelExample {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\Admin\\Desktop\\namebook.xlsx";
        String outputFilePath = "C:\\Users\\Admin\\Desktop\\namebook_swapped.xlsx";

        FileInputStream fileInputStream = null;
        Workbook workbook = null;

        try {
            fileInputStream = new FileInputStream(inputFilePath);
            workbook = new XSSFWorkbook(fileInputStream);

            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                Cell cell1 = row.getCell(0);
                Cell cell2 = row.getCell(1);

                if (cell1 != null && cell2 != null) {
                    String temp = cell1.getStringCellValue();
                    cell1.setCellValue(cell2.getStringCellValue());
                    cell2.setCellValue(temp);
                }
            }

            FileOutputStream fileOutputStream = null;

            try {
                fileOutputStream = new FileOutputStream(outputFilePath);
                workbook.write(fileOutputStream);
            } finally {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            if (fileInputStream != null) {
                try {
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
