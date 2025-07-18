package org.toys.client;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelColumnReorder {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\Admin\\Desktop\\namebook.xlsx";
        String outputFilePath = "C:\\Users\\Admin\\Desktop\\namebook_swapped.xlsx";

        FileInputStream fileInputStream = null;
        Workbook workbook = null;

        try {
            fileInputStream = new FileInputStream(inputFilePath);
            workbook = new XSSFWorkbook(fileInputStream);

            // Assuming the data is in the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Define the desired header order
            List<String> desiredOrder = Arrays.asList("Name", "Gender", "Address", "ZipCode");

            // Read the header row
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("The header row is empty.");
                return;
            }

            // Map to store the column index of each header
            Map<String, Integer> headerIndexMap = new HashMap<>();
            for (Cell cell : headerRow) {
                headerIndexMap.put(cell.getStringCellValue(), cell.getColumnIndex());
            }

            // Create a new workbook and sheet for the output
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Reordered");

            // Write the new header row in the desired order
            Row newHeaderRow = newSheet.createRow(0);
            for (int i = 0; i < desiredOrder.size(); i++) {
                Cell newCell = newHeaderRow.createCell(i);
                newCell.setCellValue(desiredOrder.get(i));
            }

            // Iterate over all rows in the original sheet and write them to the new sheet in the desired order
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowCount; i++) {
                Row originalRow = sheet.getRow(i);
                Row newRow = newSheet.createRow(i);

                for (int j = 0; j < desiredOrder.size(); j++) {
                    String header = desiredOrder.get(j);
                    Integer columnIndex = headerIndexMap.get(header);
                    if (columnIndex != null) {
                        Cell originalCell = originalRow.getCell(columnIndex);
                        if (originalCell != null) {
                            Cell newCell = newRow.createCell(j);
                            switch (originalCell.getCellType()) {
                                case STRING:
                                    newCell.setCellValue(originalCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    if (DateUtil.isCellDateFormatted(originalCell)) {
                                        newCell.setCellValue(originalCell.getDateCellValue());
                                    } else {
                                        newCell.setCellValue(originalCell.getNumericCellValue());
                                    }
                                    break;
                                case BOOLEAN:
                                    newCell.setCellValue(originalCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    newCell.setCellFormula(originalCell.getCellFormula());
                                    break;
                                default:
                                    newCell.setCellValue(originalCell.getStringCellValue());
                                    break;
                            }
                        }
                    }
                }
            }

            // Write the changes to a new file
            FileOutputStream fileOutputStream = null;
            try {
                fileOutputStream = new FileOutputStream(outputFilePath);
                newWorkbook.write(fileOutputStream);
            } finally {
                if (fileOutputStream != null) {
                    fileOutputStream.close();
                }
            }

            newWorkbook.close();

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

