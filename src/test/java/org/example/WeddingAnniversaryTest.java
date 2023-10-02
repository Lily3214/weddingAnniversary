package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.*;

public class WeddingAnniversaryTest2 {
    private static final String FILE_PATH = "Wedding Listing.xlsx";
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    @Test
    public void print_PreviousWeekDOMLists_fromSundayToSaturday() {
        List<String> expectedOutput = List.of("Name: Aacucs aacd mmimizamrmmsmh Masmhmmw, DOM: 2023-09-30, weekNumber: 39, Remarks: N/A");
        List<String> actualOutput = getPreviousWeekData(FILE_PATH);
        assertEquals(expectedOutput, actualOutput);
    }

    public List<String> getPreviousWeekData(String filePath) {
        List<String> output = new ArrayList<>();

        try (FileInputStream inputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            LocalDate previousSunday = getPreviousSunday();

            for (DayOfWeek dayOfWeek : DayOfWeek.values()) {
                LocalDate anniversaryDate = previousSunday.plusDays(dayOfWeek.getValue());

                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    Cell domCell = row.getCell(2);

                    if (domCell != null && domCell.getCellType() == CellType.NUMERIC) {
                        LocalDate domLocalDate = domCell.getLocalDateTimeCellValue().toLocalDate();

                        if (domLocalDate.isEqual(anniversaryDate)) {
                            String name = row.getCell(1).getStringCellValue();
                            String remarks = row.getCell(4).getStringCellValue().trim().isEmpty() ? "N/A" : row.getCell(4).getStringCellValue();
                            int weekNumber = (int) row.getCell(3).getNumericCellValue();
                            String formattedDate = DATE_FORMAT.format(domLocalDate);

                            String result = "Name: " + name + ", DOM: " + formattedDate + ", weekNumber: " + weekNumber + ", Remarks: " + remarks;
                            System.out.println(result);
                            output.add(result);
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return output;
    }

    private LocalDate getPreviousSunday() {
        LocalDate currentDate = LocalDate.now();
        LocalDate lastSunday = currentDate.minusDays(currentDate.getDayOfWeek().getValue());
        return lastSunday.minusDays(7);
    }
}

