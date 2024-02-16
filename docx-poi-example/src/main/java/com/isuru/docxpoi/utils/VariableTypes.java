package com.isuru.docxpoi.utils;

import lombok.Getter;

@Getter
public enum VariableTypes {
    FIRST_NAME("#firstName"),
    LAST_NAME("#lastName"),
    POSITION("#position"),
    GENDER("#gender"),
    DATE_OF_BIRTH("#birthDate"),
    ADDRESS("#address"),
    EMPLOYEE_ID("#employeeId");

    private String name;

    VariableTypes(String name) {
        this.name = name;
    }
}
