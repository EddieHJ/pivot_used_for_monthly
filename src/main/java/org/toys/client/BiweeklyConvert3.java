package org.toys.client;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class BiweeklyConvert3 {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\Admin\\Desktop\\Biweekly\\origin.xlsx";
        String outputFilePath = "C:\\Users\\Admin\\Desktop\\Biweekly\\demo.xlsx";

        // Define the desired header order
        List<String> desiredOrder = Arrays.asList("任务ID", "标题", "创建者", "执行者", "紧急程度", "影响级", "优先级",
                "事件来源", "联系人", "单量", "Jira工单", "组织1", "组织2", "事件类型1", "事件类型2", "事件类型3", "事件类型4",
                "报单时间", "完成时间", "是否完成", "借助伙伴资源", "Lead Time");

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

            // Create styles for the headers
            CellStyle headerStyle = createHeaderStyle(newWorkbook);

            // Write the new header row in the desired order
            Row newHeaderRow = newSheet.createRow(0);
            int newHeaderIndex = 0;
            for (String header : desiredOrder) {
                Cell newCell = newHeaderRow.createCell(newHeaderIndex);
                newCell.setCellValue(header);
                newCell.setCellStyle(headerStyle);
                newHeaderIndex++;
            }

            // Add filter to the header row
            newSheet.setAutoFilter(new CellRangeAddress(0, 0, 0, newHeaderIndex - 1));

            // Iterate over all rows in the original sheet
            int rowCount = sheet.getPhysicalNumberOfRows();
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            for (int i = 1; i < rowCount; i++) {
                Row originalRow = sheet.getRow(i);
                Row newRow = newSheet.createRow(i);

                Date reportDate = null;
                Date completionDate = null;
                int newColumnIndex = 0;

                for (String header : desiredOrder) {
                    switch (header) {
                        case "组织1":
                        case "组织2": {
                            Integer columnIndex = headerIndexMap.get("组织");
                            if (columnIndex != null) {
                                Cell originalCell = originalRow.getCell(columnIndex);
                                if (originalCell != null) {
                                    String cellValue = originalCell.getStringCellValue();
                                    String[] parts = cellValue.split("/");
                                    Cell newCell = newRow.createCell(newColumnIndex++);
                                    if (header.equals("组织1") && parts.length > 0) {
                                        newCell.setCellValue(parts[0]);
                                    } else if (header.equals("组织2") && parts.length > 1) {
                                        newCell.setCellValue(parts[1]);
                                    }
                                }
                            }
                            break;
                        }
                        case "事件类型1":
                        case "事件类型2":
                        case "事件类型3":
                        case "事件类型4": {
                            Integer columnIndex = headerIndexMap.get("事件类型");
                            if (columnIndex != null) {
                                Cell originalCell = originalRow.getCell(columnIndex);
                                if (originalCell != null) {
                                    String cellValue = originalCell.getStringCellValue();
                                    String[] parts = cellValue.split("/");
                                    Cell newCell = newRow.createCell(newColumnIndex++);
                                    int partIndex = Integer.parseInt(header.replace("事件类型", "")) - 1;
                                    if (partIndex < parts.length) {
                                        newCell.setCellValue(parts[partIndex]);
                                    }
                                }
                            }
                            break;
                        }
                        case "报单时间":
                            Integer reportDateIndex = headerIndexMap.get(header);
                            if (reportDateIndex != null) {
                                Cell originalCell = originalRow.getCell(reportDateIndex);
                                if (originalCell != null) {
                                    reportDate = dateFormat.parse(originalCell.getStringCellValue());
                                    newRow.createCell(newColumnIndex++).setCellValue(originalCell.getStringCellValue());
                                }
                            }
                            break;
                        case "完成时间":
                            Integer completionDateIndex = headerIndexMap.get(header);
                            if (completionDateIndex != null) {
                                Cell originalCell = originalRow.getCell(completionDateIndex);
                                if (originalCell != null && originalCell.getStringCellValue() != null && !originalCell.getStringCellValue().isEmpty()) {
                                    completionDate = dateFormat.parse(originalCell.getStringCellValue());
                                    newRow.createCell(newColumnIndex++).setCellValue(originalCell.getStringCellValue());
                                } else {
                                    // If the completion date is empty, set it as empty
                                    newRow.createCell(newColumnIndex++).setCellValue("");
                                }
                            }
                            break;
                        default:
                            Integer columnIndex = headerIndexMap.get(header);
                            if (columnIndex != null) {
                                Cell originalCell = originalRow.getCell(columnIndex);
                                if (originalCell != null) {
                                    Cell newCell = newRow.createCell(newColumnIndex++);
                                    copyCellValue(originalCell, newCell);
                                }
                            }
                            break;
                    }
                }

                // Calculate and set the Lead Time in minutes if both dates are available
                Cell leadTimeCell = newRow.createCell(newColumnIndex);
                if (reportDate != null && completionDate != null) {
                    long diffInMillis = completionDate.getTime() - reportDate.getTime();
                    long diffInMinutes = diffInMillis / (60 * 1000);
                    leadTimeCell.setCellValue(diffInMinutes);
                } else {
                    leadTimeCell.setCellValue("");
                }
            }

            // Write the changes to a new file
            try (FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath)) {
                newWorkbook.write(fileOutputStream);
            }
            newWorkbook.close();

        } catch (IOException | ParseException e) {
            e.printStackTrace();
        }
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
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
