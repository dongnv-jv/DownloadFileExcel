package com.utils;

import com.entity.Student;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.util.List;

public class CreateFile {


    public static ByteArrayInputStream listToExcelFile(List<Student> data) throws IOException {

        LocalDate date = LocalDate.parse("2022-12-25");

        Workbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Danh_sach_sinh_vien");

        Row row = sheet.createRow(0);
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // Creating header
        Cell cell = row.createCell(0);
        cell.setCellValue("STT");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(1);
        cell.setCellValue("Tên");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(2);
        cell.setCellValue("Tuổi");
        cell.setCellStyle(headerCellStyle);

        cell = row.createCell(3);
        cell.setCellValue("Ngày tạo");
        cell.setCellStyle(headerCellStyle);
        // Creating data rows for each customer
        for (int i = 0; i < data.size(); i++) {
            Row dataRow = sheet.createRow(i + 1);
            dataRow.createCell(0).setCellValue(i + 1);
            dataRow.createCell(1).setCellValue(data.get(i).getName());
            dataRow.createCell(2).setCellValue(data.get(i).getAge());
            dataRow.createCell(3).setCellValue(date.toString());
        }
        // Making size of column auto resize to fit with data
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);


        sheet.setColumnWidth(0, 1000);
        sheet.setColumnWidth(1, 8000);
        sheet.setColumnWidth(2, 8000);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return new ByteArrayInputStream(outputStream.toByteArray());

    }


}
