package com.patrick.exportexcell.controller;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.stereotype.Component;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

@Component
public class ExcelContoller {

    public boolean createExcel(List<Customer> data, ServletContext context, HttpServletRequest re
            , HttpServletResponse response) {
        String filePath = context.getRealPath("/resources/reports");
        File file = new File(filePath);
        boolean exists = new File(filePath).exists();
        if (!exists) {
            new File(filePath).mkdirs();
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(file + "/" + "customers" + ".xls");

            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet workSheet = workbook.createSheet("Customers");
            workSheet.setDefaultColumnWidth(30);

            HSSFCellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(HSSFColor.BLUE.index);
            headerCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

            HSSFRow headerRow = workSheet.createRow(0);

            HSSFCell firstName = headerRow.createCell(0);
            firstName.setCellValue("First Name");
            firstName.setCellStyle(headerCellStyle);

            HSSFCell secondName = headerRow.createCell(0);
            secondName.setCellValue("Second Name");
            secondName.setCellStyle(headerCellStyle);

            int i = 1;
            for (Customer customer : data) {
                HSSFRow bodyRow = workSheet.createRow(i);

                HSSFCellStyle bodyCellStyle = workbook.createCellStyle();
                bodyCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);

                HSSFCell firstNameValue = bodyRow.createCell(0);
                firstNameValue.setCellValue(customer.getFirstName());
                firstNameValue.setCellStyle(bodyCellStyle);

                HSSFCell lastNameValue = bodyRow.createCell(1);
                lastNameValue.setCellValue(customer.getSecondName());
                lastNameValue.setCellStyle(bodyCellStyle);
                i++;
            }

            workbook.write(outputStream);

            outputStream.flush();
            outputStream.close();
            return true;


        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }


    }





}
