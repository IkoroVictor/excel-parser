package org.javafunk.example.domain;

import org.javafunk.excelparser.annotations.ExcelField;
import org.javafunk.excelparser.annotations.ExcelObject;
import org.javafunk.excelparser.annotations.ParseType;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Value;
import lombok.experimental.FieldDefaults;

import java.math.BigDecimal;
import java.util.Date;

@Value
@AllArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
@ExcelObject(parseType = ParseType.ROW, start = 6, end = 8)
public class Student {

    @ExcelField(position = 2, validate = true ,validationType = ExcelField.ValidationType.SOFT, regex = "[0-2][1-9][0-9][0-9]")
    Long roleNumber;

    @ExcelField(position = 3)
    String name;

    @ExcelField(position = 4)
    Date dateOfBirth;

    @ExcelField(position = 5)
    String fatherName;

    @ExcelField(position = 6)
    String motherName;

    @ExcelField(position = 7)
    String address;

    @ExcelField(position = 8)
    BigDecimal totalScore;


    @SuppressWarnings("UnusedDeclaration")
    private Student() {
        this(null, null, null, null, null, null, null);
    }
}
