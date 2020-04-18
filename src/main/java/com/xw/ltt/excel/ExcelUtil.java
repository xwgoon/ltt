package com.xw.ltt.excel;

import com.sun.jna.platform.win32.Advapi32Util;
import com.sun.jna.platform.win32.WinReg;
import com.xw.ltt.Test;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

public class ExcelUtil {

    private static Path TEMP_DIR;
    private static final String EXCEL_PATH;

    static {
        EXCEL_PATH = Advapi32Util.registryGetStringValue(
                WinReg.HKEY_LOCAL_MACHINE, //HKEY
                "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe", //Key
                "Path"
        );
    }

    static class A {

    }

    static class B extends A {

    }

    public static void main(String[] args) throws Exception {
//        Class<?> clazz = B.class;
//        A a = (A) clazz.newInstance();
//        System.out.println(a);

//        System.out.println(System.getenv());

        String property = "java.io.tmpdir";

        // Get the temporary directory and print it.
        String tempDir = System.getProperty(property);
        System.out.println("OS temporary directory is " + tempDir);

    }

    private static boolean isFirstExcel = true;
    private static int mainSheetLastRowNum;

    private static Workbook createBook(Path path, String fileName) {
        System.out.println("正在合并【" + fileName + "】...");
        try (InputStream in = Files.newInputStream(path)) { //用完流后要关闭，否则后面无法删除临时文件夹
            Workbook book = WorkbookFactory.create(in);
            if (Test.sheetNum > book.getNumberOfSheets()) {
                System.out.println("您输入的表位置" + Test.sheetNum + "非法，【" + fileName + "】文件共有" + book.getNumberOfSheets() + "张表。");
                Test.isSuccess = false;
                return null;
            }
            return book;
        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static SXSSFWorkbook createMainBook(Path path, int sheetIndex, String fileName) throws IOException {
        Workbook book = createBook(path, fileName);
        if (book == null) return null;
        SXSSFWorkbook mainBook = new SXSSFWorkbook((XSSFWorkbook) book);
        SXSSFSheet mainSheet = mainBook.getSheetAt(sheetIndex);
        Sheet sheet = book.getSheetAt(sheetIndex);

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
        return mainBook;
    }

    private static Path createXlsxFile(Path xlsFile) throws Exception {
        TEMP_DIR = Files.createTempDirectory("ltt_");
        String xlsxFile = TEMP_DIR + "\\" + xlsFile.getFileName() + "x";
        String cmd = EXCEL_PATH + "excelcnv.exe -oice \"" + xlsFile + "\" \"" + xlsxFile + "\"";
        Process process = Runtime.getRuntime().exec(cmd);
        process.waitFor();
        return Paths.get(xlsxFile);
    }

    private static boolean isXls(String fileName) {
        return "xls".equals(fileName.substring(fileName.lastIndexOf('.') + 1));
    }

    public static void mergeExcelFiles(File file, List<Path> excelPaths) throws Exception {
        if (excelPaths == null || excelPaths.size() == 0) {
            System.out.println("没有原始数据。");
            Test.isSuccess = false;
            return;
        }

        SXSSFWorkbook mainBook = null;
        SXSSFSheet mainSheet = null;
        int sheetIndex = Test.sheetNum - 1;

        try {
//            for (int i = 0; i < excelPaths.size(); i++) {
//                Path path = excelPaths.get(i);
//                String fileName = path.getFileName().toString();
//                if (isXlsx(fileName)) {
//                    mainBook = createMainBook(path, sheetIndex, fileName);
//                    if (mainBook == null) return;
//                    mainBookIndex = i;
//                    break;
//                }
//            }
//
//            if (mainBook == null) {
//                mainBookIndex = 0;
//                Path xlsFile = excelPaths.get(mainBookIndex);
//                Path xlsxFile = createXlsxFile(xlsFile);
//                mainBook = createMainBook(xlsxFile, sheetIndex, xlsFile.getFileName().toString());
//                if (mainBook == null) return;
//            }

            for (int i = 0; i < excelPaths.size(); i++) {
                Path path = excelPaths.get(i);
                String fileName = path.getFileName().toString();
                if (isXls(fileName)) {
                    path = createXlsxFile(path);
                }

                if (i == 0) {
                    mainBook = createMainBook(path, sheetIndex, fileName);
                    if (mainBook == null) return;
                    mainSheet = mainBook.getSheetAt(0);
                    continue;
                }

                Workbook book = createBook(path, fileName);
                if (book == null) return;
                Sheet sheet = book.getSheetAt(sheetIndex);
                copySheets(mainBook, mainSheet, sheet);
            }

//        for (InputStream fin : list) {
//            XSSFWorkbook b = new XSSFWorkbook(fin);
////            for (int i = 0; i < b.getNumberOfSheets(); i++) {
////                copySheets(book, sheet, b.getSheetAt(i));
////            }
//
//            copySheets(book, sheet, b.getSheetAt(2));
//        }

            writeFile(mainBook, file);
        } catch (Exception e) {
            Test.isSuccess = false;
            e.printStackTrace();
        } finally {
            deleteTempDir();
        }
    }

    private static void deleteTempDir() throws IOException {
        Files.walk(TEMP_DIR)
                .map(Path::toFile)
                .sorted(Comparator.reverseOrder())
                .forEach(File::delete);
    }

    private static void writeFile(Workbook book, File file) throws Exception {
        FileOutputStream out = new FileOutputStream(file);
        book.write(out);
        out.close();
    }

    private static void copySheets(Workbook newWorkbook, Sheet newSheet, Sheet sheet) {
        copySheets(newWorkbook, newSheet, sheet, true);
    }

    private static void copySheets(Workbook newWorkbook, Sheet newSheet, Sheet sheet, boolean copyStyle) {
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
//        Set<CellRangeAddress> mergedRegions = new HashSet<>();
//        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
        for (int i = Test.titleRowNum; i <= sheet.getLastRowNum(); i++) {
            Row srcRow = sheet.getRow(i);
            if (srcRow == null || srcRow.getCell(0) == null) break;

//            if (srcRow == null) {
//                continue;
//            }
//            boolean skip = true;
//            for (int j = 0; j <= srcRow.getLastCellNum(); j++) {
//                if (hasEffectiveValue(srcRow, j)) {
//                    skip = false;
//                    break;
//                }
//            }
//            if (skip) {
//                continue;
//            }

            Row destRow = newSheet.createRow(i + newRowNum);
            copyRow(newWorkbook, srcRow, destRow, styleMap, null);
            if (srcRow.getLastCellNum() > maxColumnNum) {
                maxColumnNum = srcRow.getLastCellNum();
            }
        }
        for (int i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    private static void copyRow(Workbook newWorkbook, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap, Set<CellRangeAddress> mergedRegions) {
        destRow.setHeight(srcRow.getHeight());
        for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            Cell oldCell = srcRow.getCell(j);
            Cell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                copyCell(newWorkbook, oldCell, newCell, styleMap);

//                CellRangeAddress mergedRegion = getMergedRegion(srcRow.getSheet(), srcRow.getRowNum(),
//                        (short) oldCell.getColumnIndex());
//
//                if (mergedRegion != null) {
//                    // System.out.println("Selected merged region: " +
//                    // mergedRegion.toString());
//                    CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.getFirstRow(),
//                            mergedRegion.getLastRow(), mergedRegion.getFirstColumn(), mergedRegion.getLastColumn());
//                    // System.out.println("New merged region: " +
//                    // newMergedRegion.toString());
////                    CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(newMergedRegion);
//                    if (isNewMergedRegion(newMergedRegion, mergedRegions)) {
//                        mergedRegions.add(newMergedRegion);
//                        destRow.getSheet().addMergedRegion(newMergedRegion);
//                    }
//                }
            }
        }
    }

    private static void copyCell(Workbook newWorkbook, Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
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

//    private static boolean hasEffectiveValue(Row row, int cellNum) {
//        Cell cell = row.getCell(cellNum);
//        if (cell == null) return false;
//        cell.setCellType(CellType.STRING);
//        String value = cell.toString();
//        return !isBlank(value);
//    }
//
//    private static boolean isBlank(final CharSequence cs) {
//        int strLen;
//        if (cs == null || (strLen = cs.length()) == 0) {
//            return true;
//        }
//        for (int i = 0; i < strLen; i++) {
//            if (!Character.isWhitespace(cs.charAt(i))) {
//                return false;
//            }
//        }
//        return true;
//    }

//    private static CellRangeAddress getMergedRegion(Sheet sheet, int rowNum, short cellNum) {
//        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
//            CellRangeAddress merged = sheet.getMergedRegion(i);
//            if (merged.isInRange(rowNum, cellNum)) {
//                return merged;
//            }
//        }
//        return null;
//    }
//
//    private static boolean isNewMergedRegion(CellRangeAddress newMergedRegion, Set<CellRangeAddress> mergedRegions) {
//        return !mergedRegions.contains(newMergedRegion);
//    }


}
