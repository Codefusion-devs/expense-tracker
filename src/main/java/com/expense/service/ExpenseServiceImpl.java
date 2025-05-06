package com.expense.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.expense.model.Expense;

@Service
public class ExpenseServiceImpl implements ExpenseServiceI {

    private final String BASE_PATH = "C:\\Users\\pavan\\Desktop\\Expense\\";

    private String getUserFilePath(String username) {
        return BASE_PATH + "expenses_" + username + ".xlsx";
    }

    @Override
    public void addExpense(Expense expense, String username) throws IOException {
        String filePath = getUserFilePath(username);
        Workbook workbook;
        Sheet sheet;
        File file = new File(filePath);

        // Use current month as sheet name
        String currentMonth = LocalDateTime.now().format(DateTimeFormatter.ofPattern("MMMM yyyy"));

        if (file.exists()) {
            workbook = new XSSFWorkbook(new FileInputStream(file));
            sheet = workbook.getSheet(currentMonth);
            if (sheet == null) {
                sheet = workbook.createSheet(currentMonth);
                createHeaderRow(workbook, sheet);
            }
        } else {
            workbook = new XSSFWorkbook();
            sheet = workbook.createSheet(currentMonth);
            createHeaderRow(workbook, sheet);
        }

        int lastRow = sheet.getLastRowNum();
        Row newRow = sheet.createRow(lastRow + 1);

        String dateValue = (expense.getDate() != null) ? expense.getDate().toString() : java.time.LocalDate.now().toString();
        newRow.createCell(0).setCellValue(dateValue);
        newRow.createCell(1).setCellValue(expense.getCategory());
        newRow.createCell(2).setCellValue(expense.getAmount());
        newRow.createCell(3).setCellValue(expense.getDescription());

        updateTotalAmount(workbook, sheet);

        for (int i = 0; i < 4; i++) {
            sheet.autoSizeColumn(i);
        }

        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            workbook.write(outputStream);
        } finally {
            workbook.close();
        }
    }

    @Override
    public InputStream getExpenseFile(String username) throws IOException {
        String filePath = getUserFilePath(username);
        File file = new File(filePath);
        return Files.newInputStream(Paths.get(filePath));
    }


    private void createHeaderRow(Workbook workbook, Sheet sheet) {
        // Total Row
        Font blueFont = workbook.createFont();
        blueFont.setBold(true);
        blueFont.setColor(IndexedColors.BLUE.getIndex());
        blueFont.setFontHeightInPoints((short) 14);

        CellStyle blueLabelStyle = workbook.createCellStyle();
        blueLabelStyle.setFont(blueFont);

        Font sizeFont = workbook.createFont();
        sizeFont.setBold(true);
        sizeFont.setFontHeightInPoints((short) 14);

        CellStyle sizeFontStyle = workbook.createCellStyle();
        sizeFontStyle.setFont(sizeFont);

        Row totalRow = sheet.createRow(0);
        Cell labelCell = totalRow.createCell(1);
        labelCell.setCellValue("Total Amount");
        labelCell.setCellStyle(blueLabelStyle);

        Cell amountCell = totalRow.createCell(2);
        amountCell.setCellValue(0);
        amountCell.setCellStyle(sizeFontStyle);

        // Header Row
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.RED.getIndex());

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);

        Row header = sheet.createRow(1);
        String[] headers = { "Date", "Category", "Amount", "Description" };
        for (int i = 0; i < headers.length; i++) {
            Cell cell = header.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
            sheet.autoSizeColumn(i);
        }
    }

    private void updateTotalAmount(Workbook workbook, Sheet sheet) {
        double total = 0;
        for (int i = 2; i <= sheet.getLastRowNum(); i++) {
            Row r = sheet.getRow(i);
            if (r != null && r.getCell(2) != null) {
                total += r.getCell(2).getNumericCellValue();
            }
        }

        Row totalRow = sheet.getRow(0);
        if (totalRow == null) totalRow = sheet.createRow(0);
        Cell amountCell = totalRow.getCell(2);
        if (amountCell == null) amountCell = totalRow.createCell(2);

        amountCell.setCellValue(total);

        Font sizeFont = workbook.createFont();
        sizeFont.setBold(true);
        sizeFont.setFontHeightInPoints((short) 14);

        CellStyle style = workbook.createCellStyle();
        style.setFont(sizeFont);

        amountCell.setCellStyle(style);
    }
}
