package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Scanner;


public class WeddingAnniversary {
    private static final String DATE_FORMAT = "MM/dd/yy";
    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);
        boolean running = true;

        // Ask user for the option, menu will run until user choose to exit
        while (running) {
            System.out.println("\nWelcome to Wedding Anniversary App");
            System.out.println("Choose an option:");
            System.out.println("A) Print All Wedding Day Lists");
            System.out.println("B) Print Previous Week Wedding Anniversary from Sunday to Saturday");
            System.out.println("X) Exit");

            String input = scanner.nextLine().trim();

            switch (input.toUpperCase()) {
                case "A":
                    System.out.print("Enter the excel file name: ");
                    String filePath = scanner.nextLine().trim();
                    ExcelFileReader.readExcelFile(filePath);
                    break;
                case "B":
                    printPreviousWeekWeddingAnniversaryLists();
                    break;
                case "X":
                    running = false;
                    break;
                default:
                    System.out.println("Invalid option");
                    break;
            }
        }
        scanner.close();
    }

    // printPreviousWeekWeddingAnniversaryLists print Wedding Anniversary List for the Previous Week from Sunday to Saturday from the Excel file
    public static void printPreviousWeekWeddingAnniversaryLists() {
        // Open the Excel file using FileInputStream
        try (FileInputStream inputStream = new FileInputStream("Wedding Listing.xlsx");
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            // Get the first sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);

            LocalDate currentDate = LocalDate.now();
            // Calculate the date for the most recent Sunday
            LocalDate lastSunday = currentDate.minusDays(currentDate.getDayOfWeek().getValue());
            // Calculate the date for the previous Sunday
            LocalDate previousSunday = lastSunday.minusDays(7);

            System.out.println("Wedding Anniversary List for the Previous Week from Sunday to Saturday:");

            SimpleDateFormat dateFormat = new SimpleDateFormat(DATE_FORMAT);

            for (DayOfWeek dayOfWeek : DayOfWeek.values()) {
                // Calculate date of a specific day(ex: Monday) of the week within the previous week
                LocalDate anniversaryDate = previousSunday.plusDays(dayOfWeek.getValue());

                // Iterate through each row in Excel file
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    // Extract the third column from the current row. Which is Date of Marriage cell (DOM)
                    Cell domCell = row.getCell(2);

                    // Check if the cell is not null and if it's NUMERIC type then it converts the value into the LocalDate
                    if (domCell != null && domCell.getCellType() == CellType.NUMERIC) {
                        LocalDate domLocalDate = domCell.getLocalDateTimeCellValue().toLocalDate();

                        // Check if the LocalDate is equal to the specific anniversary date within the previous week from Sunday to Saturday
                        if (domLocalDate.isEqual(anniversaryDate)) {

                            String name = row.getCell(1).getStringCellValue();
                            String remarks = row.getCell(4).getStringCellValue();
                            int weekNumber = (int) row.getCell(3).getNumericCellValue();
                            // If remarks cell is empty, print out as N/A
                            remarks = remarks.trim().isEmpty() ? "N/A" : remarks;
                            System.out.println("Name: " + name + ", DOM: " + dateFormat.format(Date.from(domLocalDate.atStartOfDay(ZoneId.systemDefault()).toInstant())) + ", weekNumber: " + weekNumber + ", Remarks: " + remarks);
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
