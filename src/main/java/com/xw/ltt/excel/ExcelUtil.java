package com.xw.ltt.excel;

import com.xw.ltt.Test;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelUtil {

    private static boolean isFirstExcel = true;
    private static int mainSheetLastRowNum;

    public static void mergeExcelFiles(File file, List<Path> excelPaths) throws IOException {
        if (excelPaths == null || excelPaths.size() == 0) {
            System.out.println("没有原始数据。");
            Test.isSuccess = false;
            return;
        }

        SXSSFWorkbook mainBook = null;
        SXSSFSheet mainSheet = null;
        for (int i = 0; i < excelPaths.size(); i++) {
            Path path = excelPaths.get(i);
            String fileName = path.getFileName().toString();
            System.out.println("正在合并【" + fileName + "】...");
            InputStream in = Files.newInputStream(path);
            XSSFWorkbook book = new XSSFWorkbook(in);
            if (Test.sheetNum > book.getNumberOfSheets()) {
                System.out.println("您输入的表位置" + Test.sheetNum + "非法，【" + fileName + "】文件共有" + book.getNumberOfSheets() + "张表。");
                Test.isSuccess = false;
                return;
            }
            int sheetIndex = Test.sheetNum - 1;
            Sheet sheet = book.getSheetAt(sheetIndex);
            if (i == 0) {
                mainBook = new SXSSFWorkbook(book);
                mainSheet = mainBook.getSheetAt(sheetIndex);
                for (int k = 0; k < mainBook.getNumberOfSheets(); ) {
                    if (!mainBook.getSheetName(k).equals(mainSheet.getSheetName())) {
                        mainBook.removeSheetAt(k);
                    } else {
                        k++;
                    }
                }
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                    Row row = sheet.getRow(j);
                    if (row != null && row.getCell(0) == null) {
                        sheet.removeRow(row); //删除行后，sheet.getLastRowNum()的值不会变
                    }
                }
                mainSheetLastRowNum = sheet.getLastRowNum();
            } else {
                copySheets(mainBook, mainSheet, sheet);
            }
        }

//        for (InputStream fin : list) {
//            XSSFWorkbook b = new XSSFWorkbook(fin);
////            for (int i = 0; i < b.getNumberOfSheets(); i++) {
////                copySheets(book, sheet, b.getSheetAt(i));
////            }
//
//            copySheets(book, sheet, b.getSheetAt(2));
//        }

        try {
            writeFile(mainBook, file);
        } catch (Exception e) {
            Test.isSuccess = false;
            e.printStackTrace();
        }
    }

    private static void writeFile(SXSSFWorkbook book, File file) throws Exception {
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
    }

    private static void copySheets(SXSSFWorkbook newWorkbook, Sheet newSheet, Sheet sheet) {
        copySheets(newWorkbook, newSheet, sheet, true);
    }

    private static void copySheets(SXSSFWorkbook newWorkbook, Sheet newSheet, Sheet sheet, boolean copyStyle) {
        int newSheetLastRowNum;
        if (isFirstExcel) {
            newSheetLastRowNum = mainSheetLastRowNum;
            isFirstExcel = false;
        } else {
            newSheetLastRowNum = newSheet.getLastRowNum();
        }
        int newRowNum = newSheetLastRowNum + 1 - Test.titleRowNum;
        int maxColumnNum = 0;
        Map<Integer, CellStyle> styleMap = copyStyle ? new HashMap<>() : null;

//        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
        for (int i = Test.titleRowNum; i <= sheet.getLastRowNum(); i++) {
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

    private static void copyRow(SXSSFWorkbook newWorkbook, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap) {
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

    private static void copyCell(SXSSFWorkbook newWorkbook, Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
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
                newCell.setBlank();
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
