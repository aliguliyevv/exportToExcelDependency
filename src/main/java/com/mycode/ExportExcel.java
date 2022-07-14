package com.mycode;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;


public class ExportExcel<T> {
    private static final Logger logger = LogManager.getLogger(ExportExcel.class);

    public static <T> void toExcel(List<T> object, String path) throws ClassNotExportable, NoExportableFieldsFound, IOException, IllegalAccessException {
        List<String> excelElements = new ArrayList<>();
        Set<String> strings = new TreeSet<>();
        Class<?> toExcelClass = object.get(0).getClass();

        if (!toExcelClass.isAnnotationPresent(Exportable.class)) {
            logger.error("No such class with excel annotation");
            throw new ClassNotExportable("No such class with excel annotation");
        }

        int count = 0;
        for (var item : object) {
            for (Field field : item.getClass().getDeclaredFields()) {
                if (field.isAnnotationPresent(ExportFiled.class)) {
                    excelElements.add(field.get(object.get(count)).toString());
                    strings.add(field.getAnnotation(ExportFiled.class).columnName());
                }
            }
            count++;
        }
        if (count == 0) {
            logger.error("No such class filed with excel annotation");
            throw new NoExportableFieldsFound("No such class filed with excel annotation");
        }


        addToExcel(toExcelClass.getAnnotation(Exportable.class).sheetName(),strings,excelElements,path);


    }

    private static void addToExcel(String sheetName, Set<String> strings, List<String> excelElements, String path) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        // creating sheet with name "Report" in workbook
        XSSFSheet sheet = workbook.createSheet(sheetName);
        // this method creates header for our table
        Font headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setItalic(true);

        headerFont.setColor(IndexedColors.WHITE.getIndex());

        CellStyle style = workbook.createCellStyle();
        style.setFont(headerFont);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.PINK.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        Row headerRow = sheet.createRow(0);
        int cellCount = 0;
        for (var columnName : strings) {
            Cell cell = headerRow.createCell(cellCount);
            cell.setCellValue(columnName);
            cell.setCellStyle(style);
            cellCount++;
        }


        int rowCount = 1;
        int createdCell = 0;
        Font cellFont = workbook.createFont();
        cellFont.setFontHeightInPoints((short) 12);
        CellStyle styleCell = workbook.createCellStyle();
        styleCell.setFont(cellFont);
        styleCell.setBorderBottom(BorderStyle.THIN);
        styleCell.setBorderTop(BorderStyle.THIN);
        styleCell.setBorderLeft(BorderStyle.THIN);
        styleCell.setBorderRight(BorderStyle.THIN);
        styleCell.setAlignment(HorizontalAlignment.CENTER);
        Row row = sheet.createRow(rowCount);
        for (var user : excelElements) {

            Cell cell = row.createCell(createdCell);
            cell.setCellValue(user);
            cell.setCellStyle(styleCell);
            createdCell++;
            if (createdCell == cellCount) {
                rowCount++;
                row = sheet.createRow(rowCount);
                createdCell = 0;
            }
        }


        try (
                FileOutputStream outputStream = new FileOutputStream(path)) {
            workbook.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            logger.error(e.getMessage());
        } finally {
            // don't forget to close workbook to prevent memory leaks
            workbook.close();
        }
    }


}
