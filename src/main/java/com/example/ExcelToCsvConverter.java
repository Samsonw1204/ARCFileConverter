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

public class ExcelToCsvConverter {

    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.println("Welcome to the Excel to CSV Converter.");
            System.out.println("If you downloaded the Excel file from RecDesk, it may be in your Downloads folder.");
            System.out.println("Example file path: C:/Users/YourName/Downloads/RosterExtract-XXXXXX.xls");

            String inputFilePath = getFilePath(scanner);
            String outputFilePath = new File(inputFilePath).getParent() + "/BlendedClassSetup.csv";

            try {
                List<Student> students = extractStudentData(inputFilePath);
                validateData(students);
                writeCsv(students, outputFilePath);
                System.out.println("[32mCSV file created successfully at: " + outputFilePath + "[0m");
            } catch (IOException e) {
                System.err.println("[31mCSV creation failed due to program termination.[0m");
                System.err.println("[31mError: " + e.getMessage() + "[0m");
            }
        }
    }

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

    private static List<Student> extractStudentData(String filePath) throws IOException {
        List<Student> students = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath); Workbook workbook = new HSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            int firstNameIndex = Utils.getColumnIndex(sheet, "First Name");
            int lastNameIndex = Utils.getColumnIndex(sheet, "Last Name");
            int emailIndex = Utils.getColumnIndex(sheet, "Participants Email:");
            int phoneIndex = Utils.getColumnIndex(sheet, "Participants Phone:");
            int homePhoneIndex = Utils.getColumnIndex(sheet, "Home Phone"); // Fallback

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                try {
                    String firstName = Utils.getCellValue(row.getCell(firstNameIndex));
                    String lastName = Utils.getCellValue(row.getCell(lastNameIndex));
                    String email = Utils.getCellValue(row.getCell(emailIndex));
                    String parentEmail = Utils.getCellValue(row.getCell(emailIndex + 1)); // Parent email is next to participant email

                    if (!Utils.isValidEmail(email)) {
                        Utils.logSkippedRow(row.getRowNum() + 1, "\u001B[31mInvalid email: " + email + "\u001B[0m");
                        continue; // Skip row
                    }

                    String phone = Utils.getCellValue(row.getCell(phoneIndex));
                    if (!Utils.isValidPhone(phone)) {
                        phone = Utils.getCellValue(row.getCell(homePhoneIndex)); // Fallback to home phone
                        if (!Utils.isValidPhone(phone)) {
                            System.err.println("\u001B[31mCritical Error: No Valid Phone Number Found for " + firstName + " " + lastName + " -- These Errors Must be Resolved to Avoid Program Termination.\u001B[0m");
                            throw new IllegalArgumentException("No Valid Phone Number Found for " + firstName + " " + lastName + " -- These Errors Must be Resolved to Avoid Program Termination.");
                        }
                    }

                    students.add(new Student(firstName, lastName, email, phone));

                    // Pass parentEmail to validation
                    validateParentEmail(firstName, lastName, email, parentEmail);

                } catch (IllegalArgumentException e) {
                    Utils.logSkippedRow(row.getRowNum() + 1, "\u001B[31m" + e.getMessage() + "\u001B[0m");
                }
            }
        }

        return students;
    }

    private static void validateParentEmail(String firstName, String lastName, String email, String parentEmail) {
        if (email.equalsIgnoreCase(parentEmail)) {
            System.out.println("\u001B[33mWarning: Participant " + firstName + " " + lastName + " has the same email as their parent.\u001B[0m");
        }
    }

    private static void validateData(List<Student> students) {
        Map<String, List<Student>> duplicateEmails = new HashMap<>();
        List<String> errorMessages = new ArrayList<>();

        for (Student student : students) {
            duplicateEmails
                .computeIfAbsent(student.email, k -> new ArrayList<>())
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
            throw new IllegalStateException("\u001B[31mCritical Error: " + String.join("; ", errorMessages) + " -- These Errors Must be Resolved to Avoid Program Termination.\u001B[0m");
        }
    }

    private static void writeCsv(List<Student> students, String filePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            writer.write("First Name,Last Name,Email,Phone\n");
            for (Student student : students) {
                writer.write(student.toString());
                writer.newLine();
            }
        }
    }
}
