package org.toys.client;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class BiweeklyConvert2 {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\Admin\\Desktop\\Biweekly\\origin.xlsx";
        String outputFilePath = "C:\\Users\\Admin\\Desktop\\Biweekly\\demo.xlsx";

        // Define the desired header order
        List<String> desiredOrder = Arrays.asList("任务ID", "标题", "创建者", "执行者", "紧急程度", "影响级", "优先级",
                "事件来源", "联系人", "单量", "Jira工单", "组织", "事件类型", "报单时间", "完成时间", "是否完成", "借助伙伴资源");

        // Define the columns to split and the delimiter
        Map<String, List<String>> splitColumns = new HashMap<>();
        splitColumns.put("组织", Arrays.asList("组织1", "组织2"));
        splitColumns.put("事件来源", Arrays.asList("事件来源1", "事件来源2", "事件来源3", "事件来源4"));
        String delimiter = "/";

        try (FileInputStream fileInputStream = new FileInputStream(inputFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Assuming the data is in the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Read the header row and map the header indices
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                System.out.println("The header row is empty.");
                return;
            }

            // Map to store the column index of each header
            Map<String, Integer> headerIndexMap = new HashMap<>();
            for (Cell cell : headerRow) {
                headerIndexMap.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
            }

            // Create a new workbook and sheet for the output
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Reordered_Split");

            // Write the new header row in the desired order
            Row newHeaderRow = newSheet.createRow(0);
            int newHeaderIndex = 0;
            for (String header : desiredOrder) {
                if (splitColumns.containsKey(header)) {
                    for (String splitHeader : splitColumns.get(header)) {
                        newHeaderRow.createCell(newHeaderIndex++).setCellValue(splitHeader);
                    }
                } else {
                    newHeaderRow.createCell(newHeaderIndex++).setCellValue(header);
                }
            }

            // Iterate over all rows in the original sheet and write the data to the new sheet
            int rowCount = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowCount; i++) {
                Row originalRow = sheet.getRow(i);
                Row newRow = newSheet.createRow(i);

                int newColumnIndex = 0;
                for (String header : desiredOrder) {
                    Integer columnIndex = headerIndexMap.get(header);

                    if (columnIndex != null) {
                        Cell originalCell = originalRow.getCell(columnIndex);
                        if (originalCell != null) {
                            if (splitColumns.containsKey(header)) {
                                // Split the column value and create new cells
                                String cellValue = originalCell.getStringCellValue();
                                String[] parts = cellValue.split(delimiter);
                                List<String> splitHeaders = splitColumns.get(header);
                                for (int j = 0; j < splitHeaders.size(); j++) {
                                    Cell newCell = newRow.createCell(newColumnIndex++);
                                    if (j < parts.length) {
                                        newCell.setCellValue(parts[j]);
                                    } else {
                                        newCell.setCellValue(""); // Empty if not enough parts
                                    }
                                }
                            } else {
                                Cell newCell = newRow.createCell(newColumnIndex++);
                                copyCellValue(originalCell, newCell);
                            }
                        }
                    }
                }
            }

            // Write the changes to a new file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(fileOutputStream);
            }
            newWorkbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            default:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
        }
    }
}

