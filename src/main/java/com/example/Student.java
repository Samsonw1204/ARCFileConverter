package com.example;

/**
 * Represents a student with basic details such as first name, last name, email, and phone number.
 * This class is used to store and process student data extracted from the Excel file.
 */
public class Student {

    /**
     * The first name of the student.
     */
    public String firstName;

    /**
     * The last name of the student.
     */
    public String lastName;

    /**
     * The email address of the student.
     */
    public String email;

    /**
     * The phone number of the student. This field is optional and may be null or empty.
     */
    public String phone;

    /**
     * Constructs a new {@code Student} object with the specified details.
     *
     * @param firstName The first name of the student.
     * @param lastName  The last name of the student.
     * @param email     The email address of the student.
     * @param phone     The phone number of the student. This can be null or empty if not available.
     */
    public Student(String firstName, String lastName, String email, String phone) {
        this.firstName = firstName != null ? firstName.trim() : "";
        this.lastName = lastName != null ? lastName.trim() : "";
        this.email = email != null ? email.trim() : "";
        this.phone = phone != null ? phone.trim() : "";
    }

    /**
     * Returns a string representation of the student object formatted for CSV output.
     * The phone number is stripped of non-digit characters before being included.
     *
     * @return A CSV-formatted string containing the student's first name, last name, email, and phone.
     */
    @Override
    public String toString() {
        String formattedPhone = phone != null ? phone.replaceAll("\\D", "") : "";

        return String.join(",",
            escapeCsv(firstName),
            escapeCsv(lastName),
            escapeCsv(email),
            escapeCsv(formattedPhone)
        );
    }

    /**
     * Escapes values for safe CSV output.
     */
    private String escapeCsv(String value) {
        if (value == null) {
            return "";
        }

        if (value.contains(",") || value.contains("\"") || value.contains("\n") || value.contains("\r")) {
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }

        return value;
    }
}