package com.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;

public class ApachePoiBarColumnChart {
    private static final String fileName1 = "C:\\apache-poi-yasmin\\src\\main\\resources\\jira-old.xlsx";
    private static final String fileName2 = "C:\\apache-poi-yasmin\\src\\main\\resources\\jira-new.xlsx";

    public static void main(String[] args) throws IOException {
        addNewCellToExcelFile(fileName1, fileName2);
    }

    private static void addNewCellToExcelFile(String fileName1, String fileName2) {

        Workbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fileName1);
            Sheet sheet = workbook.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            Row row = sheet.createRow(lastRow + 1);

            Cell cell = row.createCell((short) 0);
            cell.setCellValue("2023 Q1 Sprint 4");

            cell = row.createCell((short) 1);
            cell.setCellValue(100);

            cell = row.createCell((short) 2);
            cell.setCellValue(200);

            try (FileOutputStream fileOut = new FileOutputStream(fileName2)) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
