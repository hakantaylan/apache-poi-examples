package com.isuru.docxpoi.controller;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
@RequestMapping("/v1")
public class ExcelController {

    @GetMapping("/download-excel")
    public ResponseEntity<byte[]> downloadExcel() throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream stream = new ByteArrayOutputStream()) {
            // Create a new Excel workbook and sheet

            Sheet sheet = workbook.createSheet("SampleSheet");
            // Create sample data (you can replace this with your own data)
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Name");
            row.createCell(1).setCellValue("Age");
            row = sheet.createRow(1);
            row.createCell(0).setCellValue("Meduim");
            row.createCell(1).setCellValue(30);
            // Write the workbook to a ByteArrayOutputStream
            workbook.write(stream);
            // Set response headers
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            headers.setContentDispositionFormData("attachment", "sample.xlsx");
            return ResponseEntity.ok()
                    .headers(headers)
                    .body(stream.toByteArray());
        }
    }

}
