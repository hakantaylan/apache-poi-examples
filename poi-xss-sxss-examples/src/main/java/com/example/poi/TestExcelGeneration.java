package com.example.poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Date;
import java.util.List;
import java.util.UUID;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestExcelGeneration {
    private static final List<Integer> rowNumberValues = List.of(10000, 50000, 100000, 200000, 500000, 1000000, 1048575);
    private static final List<Integer> streamWindowValues = List.of(100, 200, 500, 1000, 5000, 10000, 50000);
    private static final long currentMillis = System.currentTimeMillis();

    public static void main(String[] args) throws IOException {
        boolean streamMode = args.length > 0 && args[0].toLowerCase().startsWith("stream");
        // JIT Warm-up
        buildWorkbook(IntStream.rangeClosed(1, 100).boxed().collect(Collectors.toList()));
        buildWorkbook(IntStream.rangeClosed(1, 100).boxed().collect(Collectors.toList()), true, 1_000);
        if (streamMode) {
            for (int streamWindow : streamWindowValues) {
                buildWorkbook(rowNumberValues, true, streamWindow);
            }
        } else {
            buildWorkbook(rowNumberValues);
        }
    }

    public static void buildWorkbook(List<Integer> rowNumberValues, boolean streamMode, int streamWindow) throws IOException {
        System.out.println("\nBuilding workbook in " + (streamMode ? "stream mode with " + streamWindow + " row window" : "normal mode"));
        Path root = Paths.get("./tmp/excel/");
        Files.createDirectories(root);
        for (int numRows : rowNumberValues) {
            List<List<? extends Serializable>> content = IntStream.range(0, numRows).mapToObj(row -> List.of(
                    row, Math.random() * 1000000, new Date(currentMillis), UUID.randomUUID(), UUID.randomUUID(),
                    UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(),
                    UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(),
                    UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID(), UUID.randomUUID()
            )).collect(Collectors.toList());

            Workbook workbook = streamMode ? new SXSSFWorkbook(streamWindow) : new XSSFWorkbook();
            long thenMillis = System.currentTimeMillis();
            addSheet(workbook, "name_" + numRows, content);
            long nowMillis = System.currentTimeMillis();
            long elapsedMillis = nowMillis - thenMillis;
            try {
                workbook.write(new FileOutputStream("./tmp/excel/excel_" + numRows + ".xlsx"));
            } catch (Exception e) {
                e.printStackTrace();
            }
            System.out.println("Elapsed " + elapsedMillis + " millis for " + numRows + " rows");
        }
    }

    public static void buildWorkbook(List<Integer> rowNumberValues) throws IOException {
        buildWorkbook(rowNumberValues, false, 1_000);
    }

    public static void addSheet(Workbook workbook, String name, List<List<? extends Serializable>> content) {
        var wkbSheet = workbook.createSheet(name);
        int totalRows = content.size();
        int totalCells = content.get(0).size();
        for (int rowId = 0; rowId < totalRows; rowId++) {
            var row = wkbSheet.createRow(rowId);
            for (int cellId = 0; cellId < totalCells; cellId++) {
                var cell = row.createCell(cellId);
                Object value = content.get(rowId).get(cellId);
                if (value == null) {
                    cell.setBlank();
                } else if (value instanceof Integer) {
                    cell.setCellValue((Integer) value);
                } else if (value instanceof Long) {
                    cell.setCellValue((Long) value);
                } else if (value instanceof Double) {
                    cell.setCellValue((Double) value);
                } else if (value instanceof Float) {
                    cell.setCellValue((Float) value);
                } else if (value instanceof Boolean) {
                    cell.setCellValue((Boolean) value);
                } else if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else {
                    cell.setCellValue(value.toString()); // For other types just set their String representation
                }
            }
        }
    }
}

