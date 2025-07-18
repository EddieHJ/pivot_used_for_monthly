package org.toys.client;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class BiweeklyConvertPivot {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\Admin\\Desktop\\book.xlsx";
        String outputFilePath = "C:\\Users\\Admin\\Desktop\\book_reordered.xlsx";

        // Define the desired header order
        List<String> desiredOrder = Arrays.asList("任务ID", "标题", "创建者", "执行者", "紧急程度", "影响级", "优先级",
                "事件来源", "联系人", "单量", "Jira工单", "组织", "事件类型", "报单时间", "完成时间", "是否完成", "借助伙伴资源", "Lead Time");

        // Define the columns to split and the delimiter
        Map<String, List<String>> splitColumns = new HashMap<>();
        splitColumns.put("组织", Arrays.asList("组织1", "组织2"));
        splitColumns.put("事件类型", Arrays.asList("事件类型1", "事件类型2", "事件类型3", "事件类型4"));
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

            // Create styles for the headers
            CellStyle headerStyle = createHeaderStyle(newWorkbook);

            // Write the new header row in the desired order
            Row newHeaderRow = newSheet.createRow(0);
            int newHeaderIndex = 0;
            for (String header : desiredOrder) {
                if (splitColumns.containsKey(header)) {
                    for (String splitHeader : splitColumns.get(header)) {
                        Cell newCell = newHeaderRow.createCell(newHeaderIndex);
                        newCell.setCellValue(splitHeader);
                        newCell.setCellStyle(headerStyle);
                        newHeaderIndex++;
                    }
                } else {
                    Cell newCell = newHeaderRow.createCell(newHeaderIndex);
                    newCell.setCellValue(header);
                    newCell.setCellStyle(headerStyle);
                    newHeaderIndex++;
                }
            }

            // Add filter to the header row
            newSheet.setAutoFilter(new CellRangeAddress(0, 0, 0, newHeaderIndex - 1));

            // Iterate over all rows in the original sheet
            int rowCount = sheet.getPhysicalNumberOfRows();
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            for (int i = 1; i < rowCount; i++) {
                Row originalRow = sheet.getRow(i);
                Row newRow = newSheet.createRow(i);

                int newColumnIndex = 0;
                Date reportDate = null;
                Date completionDate = null;
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
                                    Cell newCell = newRow.createCell(newColumnIndex);
                                    if (j < parts.length) {
                                        newCell.setCellValue(parts[j]);
                                    }
                                    newColumnIndex++;
                                }
                            } else {
                                Cell newCell = newRow.createCell(newColumnIndex);
                                copyCellValue(originalCell, newCell);
                                if ("报单时间".equals(header)) {
                                    reportDate = dateFormat.parse(originalCell.getStringCellValue());
                                } else if ("完成时间".equals(header)) {
                                    completionDate = dateFormat.parse(originalCell.getStringCellValue());
                                }
                                newColumnIndex++;
                            }
                        }
                    }
                }

                // Calculate and set the Lead Time in minutes
                if (reportDate != null && completionDate != null) {
                    Cell leadTimeCell = newRow.createCell(newColumnIndex);
                    long diffInMillis = completionDate.getTime() - reportDate.getTime();
                    long diffInMinutes = diffInMillis / (60 * 1000);
                    leadTimeCell.setCellValue(diffInMinutes);
                }
            }

            // Create a new sheet for the PivotTable
            Sheet pivotSheet = newWorkbook.createSheet("数据分析");

            // Create the PivotTable
            XSSFPivotTable pivotTable = ((XSSFSheet) pivotSheet).createPivotTable(
                    new AreaReference("A1:" + CellReference.convertNumToColString(newHeaderIndex - 1) + rowCount, SpreadsheetVersion.EXCEL2007),
                    new CellReference("A1"),
                    newSheet);

            // Configure the PivotTable
            // Add row labels
            pivotTable.addRowLabel(13); // Assuming you want "标题" as row label (second column)
            pivotTable.addRowLabel(14); // Assuming you want "创建者" as row label (third column)
            pivotTable.addRowLabel(15); // Assuming you want "执行者" as row label (fourth column)
            pivotTable.addRowLabel(16);

            // Add column labels
            //pivotTable.addColLabel(4); // Assuming you want "紧急程度" as column label (fifth column)

            // Add value columns
            pivotTable.addColumnLabel(DataConsolidateFunction.COUNT, 2, "count");




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
