package com.xw.ltt.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    public static void mergeExcelFiles(File file, List<InputStream> list) throws IOException {
        Path titleTemplatePath = Paths.get("表头模板/卡片数据.xlsx");
        InputStream inputStream = Files.newInputStream(titleTemplatePath);
        XSSFWorkbook book = new XSSFWorkbook(inputStream);
        Sheet sheet = book.getSheetAt(0);

        for (InputStream fin : list) {
            XSSFWorkbook b = new XSSFWorkbook(fin);
//            for (int i = 0; i < b.getNumberOfSheets(); i++) {
//                copySheets(book, sheet, b.getSheetAt(i));
//            }

            copySheets(book, sheet, b.getSheetAt(2));
        }

        try {
            writeFile(book, file);
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    protected static void writeFile(XSSFWorkbook book, File file) throws Exception {
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
    }

    private static void copySheets(XSSFWorkbook newWorkbook, Sheet newSheet, Sheet sheet) {
        copySheets(newWorkbook, newSheet, sheet, true);
    }

    private static void copySheets(XSSFWorkbook newWorkbook, Sheet newSheet, Sheet sheet, boolean copyStyle) {
        int newRowNum = newSheet.getLastRowNum() - 3; //i的初始值为4
        int maxColumnNum = 0;
        Map<Integer, CellStyle> styleMap = (copyStyle) ? new HashMap<Integer, CellStyle>() : null;

//        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
        for (int i = 4; i <= sheet.getLastRowNum(); i++) {
            Row srcRow = sheet.getRow(i);
            if (srcRow == null || srcRow.getCell(0) == null) break;
            Row destRow = newSheet.createRow(i + newRowNum);
            copyRow(newWorkbook, srcRow, destRow, styleMap);
            if (srcRow.getLastCellNum() > maxColumnNum) {
                maxColumnNum = srcRow.getLastCellNum();
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    public static void copyRow(XSSFWorkbook newWorkbook, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) {
        destRow.setHeight(srcRow.getHeight());
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            Cell oldCell = srcRow.getCell(j);
            Cell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                copyCell(newWorkbook, oldCell, newCell, styleMap);
            }
        }
    }

    public static void copyCell(XSSFWorkbook newWorkbook, Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
        if (styleMap != null) {
            int stHashCode = oldCell.getCellStyle().hashCode();
            CellStyle newCellStyle = styleMap.get(stHashCode);
            if (newCellStyle == null) {
                newCellStyle = newWorkbook.createCellStyle();
                newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
                styleMap.put(stHashCode, newCellStyle);
            }
            newCell.setCellStyle(newCellStyle);
        }


        CellType oldCellType = oldCell.getCellType();
        if (oldCellType == CellType.FORMULA) {
            newCell.setCellFormula(oldCell.getCellFormula());
            oldCellType = oldCell.getCachedFormulaResultType();
        }
        switch (oldCellType) {
            case STRING:
                newCell.setCellValue(oldCell.getRichStringCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case BLANK:
                newCell.setCellType(CellType.BLANK);
                break;
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case ERROR:
                newCell.setCellErrorValue(oldCell.getErrorCellValue());
                break;
//            case FORMULA:
//                newCell.setCellFormula(oldCell.getCellFormula());
//                break;
            default:
                break;
        }
    }

}
