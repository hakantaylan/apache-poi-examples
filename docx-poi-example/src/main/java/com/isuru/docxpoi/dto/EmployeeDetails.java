package com.isuru.docxpoi.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class EmployeeDetails {
    private String firstName;
    private String lastName;
    private String position;
    private String gender;
    private LocalDate dob;
    private String address;
    private Integer employeeId;
}
