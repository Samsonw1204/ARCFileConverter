package com.example;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class Utils {

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

    public static boolean isValidEmail(String email) {
        return email != null && email.contains("@") && email.contains(".");
    }

    public static boolean isValidPhone(String phone) {
        if (phone == null) return false;
        String digitsOnly = phone.replaceAll("\\D", ""); // Remove all non-numeric characters
        return digitsOnly.length() == 10;
    }

    public static void logSkippedRow(int rowNum, String message) {
        try (BufferedWriter logWriter = new BufferedWriter(new FileWriter("skipped_rows.log", true))) {
            logWriter.write("Skipped row " + rowNum + ": " + message + "\n");
        } catch (IOException e) {
            System.err.println("Failed to log skipped row " + rowNum + ": " + e.getMessage());
        }
    }
}
