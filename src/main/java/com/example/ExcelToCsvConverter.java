package com.example;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.stream.Collectors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Converts an Excel (.xls) roster file into a CSV file formatted for Red Cross
 * blended learning class setup.
 *
 * The output file contains:
 * First Name, Last Name, Email, Phone
 *
 * This version supports both:
 * - Lifeguarding roster email column:
 *   "Participants Email:"
 *
 * - WSI roster email column:
 *   "Participants Email (Recommended to NOT use a school or work email)"
 */
public class ExcelToCsvConverter {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        try {
            System.out.println("Welcome to the Excel to CSV Converter.");
            System.out.println("If you downloaded the Excel file from RecDesk, it may be in your Downloads folder.");
            System.out.println("Example file path: C:/Users/YourName/Downloads/RosterExtract-XXXXXX.xls");

            String inputFilePath = getFilePath(scanner);
            String outputFilePath = new File(inputFilePath).getParent() + "/BlendedClassSetup.csv";

            try {
                File outputFile = new File(outputFilePath);

                // Remove any old output file before processing so a failed run does not
                // leave behind a stale CSV that appears to be newly created.
                if (outputFile.exists() && !outputFile.delete()) {
                    throw new IOException("Unable to delete existing output file: " + outputFilePath);
                }

                List<Student> students = extractStudentData(inputFilePath);
                validateData(students);
                writeCsv(students, outputFilePath);

                System.out.println("\u001B[32mCSV file created successfully at: " + outputFilePath + "\u001B[0m");

            } catch (IOException e) {
                System.err.println("\u001B[31mError: CSV Creation Failed Due to Program Termination\u001B[0m");
                System.err.println("\u001B[31mDetails: " + e.getMessage() + "\u001B[0m");

            } catch (IllegalStateException e) {
                System.err.println(e.getMessage());

            } catch (Exception e) {
                System.err.println("\u001B[31mAn unexpected error occurred: " + e.getMessage() + "\u001B[0m");
            }

        } finally {
            System.out.println("Press Enter to exit...");
            scanner.nextLine();
            scanner.close();
        }
    }

    /**
     * Prompts the user for the Excel file path and validates the input.
     */
    private static String getFilePath(Scanner scanner) {
        while (true) {
            System.out.print("Enter the path to the Excel file (or type QUIT to exit): ");
            String inputFilePath = scanner.nextLine().trim();

            if (inputFilePath.equalsIgnoreCase("QUIT")) {
                System.out.println("Program terminated by user. Exiting...");
                System.exit(0);
            }

            if (inputFilePath.startsWith("\"") && inputFilePath.endsWith("\"")) {
                inputFilePath = inputFilePath.substring(1, inputFilePath.length() - 1);
            }

            File inputFile = new File(inputFilePath);

            if (!inputFile.exists() || !inputFilePath.toLowerCase().endsWith(".xls")) {
                System.out.println("Invalid file path. Please check the path and try again.");
                continue;
            }

            return inputFilePath;
        }
    }

    /**
     * Extracts student data from the Excel file.
     */
    private static List<Student> extractStudentData(String filePath) throws IOException {
        List<Student> students = new ArrayList<>();
        List<String> validationErrors = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new HSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            int firstNameIndex = Utils.getColumnIndex(sheet, "First Name");
            int lastNameIndex = Utils.getColumnIndex(sheet, "Last Name");

            int emailIndex = Utils.getOptionalColumnIndex(sheet, "Participants Email:");

            int wsiEmailIndex = Utils.getOptionalColumnIndex(
                sheet,
                "Participants Email (Recommended to NOT use a school or work email)"
            );

            int parentEmailIndex = Utils.getOptionalColumnIndex(
                sheet,
                "Parent/Guardian Email (If participant is under 18):"
            );

            int phoneIndex = Utils.getOptionalColumnIndex(sheet, "Participants Phone:");
            int homePhoneIndex = Utils.getOptionalColumnIndex(sheet, "Home Phone");

            if (emailIndex == -1 && wsiEmailIndex == -1) {
                throw new IllegalStateException(
                    "\u001B[31mCritical Error: No usable participant email column was found. "
                    + "Expected either 'Participants Email:' or "
                    + "'Participants Email (Recommended to NOT use a school or work email)'.\u001B[0m"
                );
            }

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                try {
                    String firstName = Utils.getCellValue(row.getCell(firstNameIndex));
                    String lastName = Utils.getCellValue(row.getCell(lastNameIndex));

                    if (firstName.isBlank() && lastName.isBlank()) {
                        continue;
                    }

                    String email = getBestEmail(row, emailIndex, wsiEmailIndex);

                    if (!Utils.isValidEmail(email)) {
                        String message = "Row " + (row.getRowNum() + 1)
                            + " (" + firstName + " " + lastName + ") has an invalid participant email: \""
                            + email + "\"";

                        validationErrors.add(message);
                        Utils.logSkippedRow(row.getRowNum() + 1, message);
                        continue;
                    }

                    String phone = getBestPhone(row, phoneIndex, homePhoneIndex);

                    students.add(new Student(firstName, lastName, email, phone));

                    if (parentEmailIndex != -1) {
                        String parentEmail = Utils.getCellValue(row.getCell(parentEmailIndex));
                        validateParentEmail(firstName, lastName, email, parentEmail);
                    }

                } catch (IllegalArgumentException e) {
                    String message = "Row " + (row.getRowNum() + 1) + ": " + e.getMessage();
                    validationErrors.add(message);
                    Utils.logSkippedRow(row.getRowNum() + 1, message);
                }
            }
        }

        if (!validationErrors.isEmpty()) {
            throw new IllegalStateException(
                "\u001B[31mCritical Error: One or more rows contain invalid data. "
                + "CSV file was not created. Please correct the following:\n"
                + String.join("\n", validationErrors)
                + "\u001B[0m"
            );
        }

        return students;
    }

    /**
     * Gets the best available email for the participant.
     *
     * Priority:
     * 1. Lifeguarding email column, if present and valid.
     * 2. WSI email column, if present and valid.
     */
    private static String getBestEmail(Row row, int emailIndex, int wsiEmailIndex) {
        String email = "";

        if (emailIndex != -1) {
            email = Utils.getCellValue(row.getCell(emailIndex));
        }

        if (!Utils.isValidEmail(email) && wsiEmailIndex != -1) {
            email = Utils.getCellValue(row.getCell(wsiEmailIndex));
        }

        return email;
    }

    /**
     * Gets the best available phone number.
     *
     * Phone is optional for blended class setup, so the program does not terminate
     * if no valid phone number is found.
     */
    private static String getBestPhone(Row row, int phoneIndex, int homePhoneIndex) {
        String phone = "";

        if (phoneIndex != -1) {
            phone = Utils.getCellValue(row.getCell(phoneIndex));
        }

        if (!Utils.isValidPhone(phone) && homePhoneIndex != -1) {
            phone = Utils.getCellValue(row.getCell(homePhoneIndex));
        }

        if (!Utils.isValidPhone(phone)) {
            phone = "";
        }

        return phone;
    }

    /**
     * Warns if the parent email matches the student's email.
     */
    private static void validateParentEmail(String firstName, String lastName, String email, String parentEmail) {
        if (Utils.isValidEmail(parentEmail) && email.equalsIgnoreCase(parentEmail)) {
            System.out.println(
                "\u001B[33mWarning: Participant " + firstName + " " + lastName
                + " has the same email as their parent.\u001B[0m"
            );
        }
    }

    /**
     * Validates the extracted student data for duplicate emails.
     */
    private static void validateData(List<Student> students) {
        Map<String, List<Student>> duplicateEmails = new HashMap<>();
        List<String> errorMessages = new ArrayList<>();

        for (Student student : students) {
            duplicateEmails
                .computeIfAbsent(student.email.toLowerCase(), k -> new ArrayList<>())
                .add(student);
        }

        duplicateEmails.values().stream()
            .filter(duplicates -> duplicates.size() > 1)
            .forEach(duplicates -> {
                String names = duplicates.stream()
                    .map(s -> s.firstName + " " + s.lastName)
                    .collect(Collectors.joining(" and "));

                errorMessages.add("Duplicate email found for " + names);
            });

        if (!errorMessages.isEmpty()) {
            throw new IllegalStateException(
                "\u001B[31mCritical Error: " + String.join("; ", errorMessages)
                + " -- These Errors Must be Resolved to Avoid Program Termination.\u001B[0m"
            );
        }
    }

    /**
     * Writes the student data to a CSV file.
     */
    private static void writeCsv(List<Student> students, String filePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            writer.write("First Name,Last Name,Email,Phone");
            writer.newLine();

            for (Student student : students) {
                writer.write(student.toString());
                writer.newLine();
            }
        }
    }
}