package com.example;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Utility class providing helper methods for processing Excel files and validating data.
 */
public class Utils {

    /**
     * Retrieves the value from a cell in a formatted string form, handling different cell types.
     *
     * @param cell The {@link Cell} object to extract the value from.
     * @return The string representation of the cell's value, or an empty string if the cell is null or has no valid data.
     */
    public static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING -> {
                return cell.getStringCellValue().trim();
            }
            case NUMERIC -> {
                return String.valueOf((long) cell.getNumericCellValue());
            }
            case BOOLEAN -> {
                return String.valueOf(cell.getBooleanCellValue());
            }
            case FORMULA -> {
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            }
            default -> {
                return "";
            }
        }
    }

    /**
     * Finds the index of a column in an Excel sheet by its header name.
     *
     * @param sheet      The {@link Sheet} object to search in.
     * @param columnName The name of the column header to find.
     * @return The zero-based index of the column.
     * @throws IllegalArgumentException If the header row is missing or the column name is not found.
     */
    public static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0); // Assuming the first row contains headers
        if (headerRow == null) throw new IllegalArgumentException("Header row is missing.");

        for (Cell cell : headerRow) {
            if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        throw new IllegalArgumentException("Column '" + columnName + "' not found.");
    }

    /**
     * Validates if an email address is in a valid format.
     *
     * @param email The email address to validate.
     * @return {@code true} if the email contains "@" and ".", otherwise {@code false}.
     */
    public static boolean isValidEmail(String email) {
        return email != null && email.contains("@") && email.contains(".");
    }

    /**
     * Validates if a phone number is valid.
     * A valid phone number contains exactly 10 numeric characters after removing non-numeric characters.
     *
     * @param phone The phone number to validate.
     * @return {@code true} if the phone number is valid, otherwise {@code false}.
     */
    public static boolean isValidPhone(String phone) {
        if (phone == null) return false;
        String digitsOnly = phone.replaceAll("\\D", ""); // Remove all non-numeric characters
        return digitsOnly.length() == 10;
    }

    /**
     * Logs a skipped row to a log file with the row number and a message describing the reason.
     *
     * @param rowNum  The number of the row that was skipped.
     * @param message The reason why the row was skipped.
     */
    public static void logSkippedRow(int rowNum, String message) {
        try (BufferedWriter logWriter = new BufferedWriter(new FileWriter("skipped_rows.log", true))) {
            logWriter.write("Skipped row " + rowNum + ": " + message + "\n");
        } catch (IOException e) {
            System.err.println("Failed to log skipped row " + rowNum + ": " + e.getMessage());
        }
    }
}
