package com.example;

public class Student {
    public String firstName;
    public String lastName;
    public String email;
    public String phone;

    public Student(String firstName, String lastName, String email, String phone) {
        this.firstName = firstName;
        this.lastName = lastName;
        this.email = email;
        this.phone = phone;
    }

    @Override
    public String toString() {
        String formattedPhone = phone != null ? phone.replaceAll("\\D", "") : "";
        return String.join(",", firstName, lastName, email, formattedPhone);
    }
}
