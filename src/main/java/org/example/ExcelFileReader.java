package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelFileReader {
    private static final String DATE_FORMAT = "MM/dd/yy";

    public static void readExcelFile(String filePath) {
        try (FileInputStream inputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // iterate through each row
            for (Row row : sheet) {
                String formattedRow = formatRow(row);
                if (formattedRow != null) {
                    System.out.println(formattedRow);
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String formatRow(Row row) {
        // declare column index No, Name, DOM, WeekNumber, Remarks
        int noColumnIndex = 0;
        int nameColumnIndex = 1;
        int domColumnIndex = 2;
        int weekNumberColumnIndex = 3;
        int remarksColumnIndex = 4;

        // Extract cell value from the row
        Cell noCell = row.getCell(noColumnIndex);
        Cell nameCell = row.getCell(nameColumnIndex);
        Cell domCell = row.getCell(domColumnIndex);
        Cell weekNumberCell = row.getCell(weekNumberColumnIndex);
        Cell remarksCell = row.getCell(remarksColumnIndex);

        // Check if cells are blank or contain "N/A" if yes, return null, so it skip that row
        if (areAllCellsNAOrBlank(noCell, nameCell, domCell, weekNumberCell, remarksCell)) {
            return null;
        }
        // Extract value if noCell is not null, and convert it to a string. Otherwise, it assigns the string N/A to no value
        String noValue = (noCell != null) ? String.valueOf((int) noCell.getNumericCellValue()) : "N/A";
        String nameValue = (nameCell != null) ? nameCell.getStringCellValue() : "N/A";
        String domValue = (domCell != null) ? formatDate(domCell.getDateCellValue()) : "N/A";
        String weekNumberValue = (weekNumberCell != null) ? String.valueOf((int) weekNumberCell.getNumericCellValue()) : "N/A";
        String remarksValue = (remarksCell != null) ? remarksCell.getStringCellValue() : "N/A";

        return String.format("No: %s, name: %s, DOM: %s, weekNumber: %s, Remarks: %s",
                noValue, nameValue, domValue, weekNumberValue, remarksValue);
    }

    // Format Date into SimpleDateFormat
    private static String formatDate(Date date) {
        SimpleDateFormat dateFormat = new SimpleDateFormat(DATE_FORMAT);
        return dateFormat.format(date);
    }

    // Check if all the cell are blank or contain N/A. if cell contains blank or N/A return false otherwise return true
    private static boolean areAllCellsNAOrBlank(Cell... cells) {
        for (Cell cell : cells) {
            if (cell != null) {
                if (cell.getCellType() != CellType.BLANK && cell.getCellType() != CellType._NONE) {
                    return false;
                }
            }
        }
        return true;
    }
}
