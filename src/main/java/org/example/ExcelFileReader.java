package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class WeddingAnniversary {
    public static void main(String[] args) {
        String inputFile = "Wedding Listing.xlsx"; // Replace with the path to your Excel file

        try (FileInputStream inputStream = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming you want to read the first sheet

            for (Row row : sheet) {
                // Assuming specific column indices for the data
                int noColumnIndex = 0;
                int nameColumnIndex = 1;
                int domColumnIndex = 2;
                int weekNumberColumnIndex = 3;
                int remarksColumnIndex = 4;

                Cell noCell = row.getCell(noColumnIndex);
                Cell nameCell = row.getCell(nameColumnIndex);
                Cell domCell = row.getCell(domColumnIndex);
                Cell weekNumberCell = row.getCell(weekNumberColumnIndex);
                Cell remarksCell = row.getCell(remarksColumnIndex);

                // Check if all cells are "N/A" or blank, and skip the row if true
                if (areAllCellsNAOrBlank(noCell, nameCell, domCell, weekNumberCell, remarksCell)) {
                    continue;
                }

                String noValue = (noCell != null) ? String.valueOf((int) noCell.getNumericCellValue()) : "N/A";
                String nameValue = (nameCell != null) ? nameCell.getStringCellValue() : "N/A";
                String domValue = (domCell != null) ? formatDate(domCell.getDateCellValue()) : "N/A";
                String weekNumberValue = (weekNumberCell != null) ? String.valueOf((int) weekNumberCell.getNumericCellValue()) : "N/A";
                String remarksValue = (remarksCell != null) ? remarksCell.getStringCellValue() : "N/A";

                String formattedRow = String.format("No: %s, name: %s, DOM: %s, weekNumber: %s, Remarks: %s",
                        noValue, nameValue, domValue, weekNumberValue, remarksValue);

                System.out.println(formattedRow);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String formatDate(Date date) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yy");
        return dateFormat.format(date);
    }

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

/*
        Scanner scanner = new Scanner(System.in);
        boolean running = true;

        while (running) {
            System.out.println("Welcome to Marriage Anniversary App");
            System.out.println("Choose an option:");
            System.out.println("A) Print previous week from Sunday to Saturday Marriage Anniversary");
            System.out.println("B) Print previous week from Sunday to Saturday Marriage Anniversary");
            System.out.println("X) Exit");

            String input = scanner.nextLine().trim();

            switch (input.toUpperCase()) {
                case "A":
                    printAllWeddingAnniversaryLists();
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
    public static void printAllWeddingAnniversaryLists(String fileName) {
        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            int number = 1;
            while ((line = reader.readLine()) != null) {
                String[] parts = line.split(",");
                if (parts.length >= 4) {
                    LocalDate date = LocalDate.parse(parts[2], DateTimeFormatter.ofPattern(DATE_FORMAT));
                    number++;
                    String name = parts[1];
                    int weekNumber = Integer.parseInt(parts[3]);
                    weddingDates.add(new WeddingDate(number, name, date, weekNumber));
                }
            }
        } catch (IOException | DateTimeParseException e) {
            System.out.println("Error loading birthday data: " + e.getMessage());
        }
    }
}
*/
