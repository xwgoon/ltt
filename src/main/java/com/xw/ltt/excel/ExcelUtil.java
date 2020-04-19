package com.xw.ltt.excel;

import com.sun.jna.platform.win32.Advapi32Util;
import com.sun.jna.platform.win32.WinReg;
import com.xw.ltt.Test;
import com.xw.ltt.vo.Sum;
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
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.temporal.ChronoUnit;
import java.util.*;

public class ExcelUtil {

    private static Path tmpDir;
    private static String excelCnvDir;
    private static SXSSFWorkbook mainBook = null;
    private static SXSSFSheet mainSheet = null;

    static class A {

    }

    static class B extends A {

    }

    public static void main(String[] args) throws Exception {
//        Class<?> clazz = B.class;
//        A a = (A) clazz.newInstance();
//        System.out.println(a);

//        System.out.println(System.getenv());

//        String property = "java.io.tmpdir";

        // Get the temporary directory and print it.
//        String tempDir = System.getProperty(property);
//        System.out.println("OS temporary directory is " + tempDir);

//        double d = 0.1 + 0.2;
//        System.out.println(d);
//        System.out.println(String.valueOf(d));
//        System.out.println(Double.toString(d));
//        System.out.println(new BigDecimal(Double.toString(d)));


        System.out.println(ChronoUnit.DAYS.between(LocalDate.of(2003, 4, 1), LocalDate.of(1900, 1, 1)));
    }

    private static boolean isFirstExcel = true;
    private static int mainSheetLastRowNum;
    private static String[] cellValArr = new String[30];

    private static Workbook createBook(Path path, String fileName) {
        System.out.println("正在合并【" + fileName + "】...");
        try (InputStream in = Files.newInputStream(path)) { //用完流后要关闭，否则后面无法删除临时文件夹
            Workbook book = WorkbookFactory.create(in);
            if (Test.sheetNum > book.getNumberOfSheets()) {
                System.out.println("您输入的表位置" + Test.sheetNum + "非法，【" + fileName + "】文件共有"
                        + book.getNumberOfSheets() + "张表。");
                Test.isSuccess = false;
                return null;
            }
            return book;
        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
        return null;
    }

    private static SXSSFWorkbook createMainBook(Path path, int sheetIndex, String fileName) {
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

        for (int j = Test.titleRowNum; j <= sheet.getLastRowNum(); j++) {
            Row row = sheet.getRow(j);
            if (row != null) {
                if (row.getCell(0) == null) {
                    sheet.removeRow(row); //删除行后，sheet.getLastRowNum()的值不会变
                }
//                else {
//                    for (short k = 0; k <= row.getLastCellNum(); k++) {
//                        Cell cell = row.getCell(k);
//                        cellValArr[k] = getCellVal(cell);
//                    }
//                    calcVal();
//                }
            }
        }

        mainSheetLastRowNum = sheet.getLastRowNum();
        return mainBook;
    }

    private static Path createXlsxFile(Path xlsFile) throws Exception {
        if (excelCnvDir == null) {
            excelCnvDir = Advapi32Util.registryGetStringValue(
                    WinReg.HKEY_LOCAL_MACHINE, //HKEY
                    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe", //Key
                    "Path"
            );
        }
        tmpDir = Files.createTempDirectory("ltt_");
        String xlsxFile = tmpDir + "\\" + xlsFile.getFileName() + "x";
        String cmd = excelCnvDir + "excelcnv.exe -oice \"" + xlsFile + "\" \"" + xlsxFile + "\"";
        Process process = Runtime.getRuntime().exec(cmd);
        process.waitFor();
        return Paths.get(xlsxFile);
    }

    public static void mergeExcelFiles(File file, List<Path> excelPaths) throws Exception {
        if (excelPaths == null || excelPaths.size() == 0) {
            System.out.println("没有原始数据。");
            Test.isSuccess = false;
            return;
        }

        if (Test.isCard) {
            Path templatePath = Paths.get(Test.WORK_DIR + "模板/模板.xlsx");
            try (InputStream in = Files.newInputStream(templatePath)) { //用完流后要关闭，否则后面无法删除临时文件夹
                XSSFWorkbook book = new XSSFWorkbook(in);
                mainBook = new SXSSFWorkbook(book);
                mainSheet = mainBook.getSheetAt(2);
                mainSheetLastRowNum = 3;
            }
        }

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
                if (fileName.endsWith("xls")) {
                    path = createXlsxFile(path);
                }

                if (!Test.isCard && i == 0) {
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

            fillSumSheet();

            writeFile(mainBook, file);
        } catch (Exception e) {
            Test.isSuccess = false;
            e.printStackTrace();
        } finally {
            deleteTempDir();
        }
    }

    private static void deleteTempDir() throws IOException {
        if (tmpDir != null) {
            Files.walk(tmpDir)
                    .map(Path::toFile)
                    .sorted(Comparator.reverseOrder())
                    .forEach(File::delete);
        }
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
        short maxColumnNum = 0;
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
        for (short i = 0; i <= maxColumnNum; i++) {
            newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
        }
    }

    private static void copyRow(Workbook newWorkbook, Row srcRow, Row destRow, Map<Integer, CellStyle> styleMap,
                                Set<CellRangeAddress> mergedRegions) {
        destRow.setHeight(srcRow.getHeight());
        for (short j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
            Cell oldCell = srcRow.getCell(j);
            Cell newCell = destRow.getCell(j);
            if (oldCell != null) {
                if (newCell == null) {
                    newCell = destRow.createCell(j);
                }
                cellValArr[j] = copyCell(newWorkbook, oldCell, newCell, styleMap);
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
        calcVal();
    }

    private static void calcVal() {
        String val;

        //资产类别：输电线路,变电设备,配电线路,配电设备-其他,配电设备-电动汽车充换电设备,用电计量设备,通信线路及设备,
        // 自动化控制设备、信息设备及仪器仪表,发电及供热设备,水工机械设备,制造及检修维护设备,生产管理用工器具,运输设备,辅助生产用设备及器具,
        // 房屋,建筑物,土地
        val = cellValArr[1];
        boolean eq输电线路 = "输电线路".equals(val);
        boolean eq变电设备 = "变电设备".equals(val);
        boolean eq配电线路 = "配电线路".equals(val);
        boolean eq配电设备其他 = "配电设备-其他".equals(val);
        boolean eq配电设备电动汽车充换电设备 = "配电设备-电动汽车充换电设备".equals(val);
        boolean eq用电计量设备 = "用电计量设备".equals(val);
        boolean eq通信线路及设备 = "通信线路及设备".equals(cellValArr[1]);
        boolean eq自动化控制设备信息设备及仪器仪表 = "自动化控制设备、信息设备及仪器仪表".equals(val);
        boolean eq发电及供热设备 = "发电及供热设备".equals(val);
        boolean eq水工机械设备 = "水工机械设备".equals(val);
        boolean eq制造及检修维护设备 = "制造及检修维护设备".equals(val);
        boolean eq生产管理用工器具 = "生产管理用工器具".equals(val);
        boolean eq运输设备 = "运输设备".equals(val);
        boolean eq辅助生产用设备及器具 = "辅助生产用设备及器具".equals(val);
        boolean eq房屋 = "房屋".equals(val);
        boolean eq建筑物 = "建筑物".equals(val);
        boolean eq土地 = "土地".equals(val);

        //电压等级：500kV,220kV,110kV,35kV,10kV,10kV以下
        val = cellValArr[3];
        boolean eq500kV = "500kV".equals(val);
        boolean eq220kV = "220kV".equals(val);
        boolean eq110kV = "110kV".equals(val);
        boolean eq35kV = "35kV".equals(val);
        boolean eq10kV = "10kV".equals(val);
        boolean eq10kV以下 = "10kV以下".equals(val);

        //资本化日期（2014-12-31，poi获取到的值是42004.0）
        boolean le20141231 = "42004.0".compareTo(cellValArr[4]) >= 0;

        String gVal = cellValArr[6];
        if (eq输电线路) {
            if (eq500kV) {
                if (le20141231) {
                    Sum.c6 = sum(Sum.c6, gVal);
                } else {
                    Sum.d6 = sum(Sum.d6, gVal);
                }
            } else if (eq220kV) {
                if (le20141231) {
                    Sum.c7 = sum(Sum.c7, gVal);
                } else {
                    Sum.d7 = sum(Sum.d7, gVal);
                }
            } else if (eq110kV) {
                if (le20141231) {
                    Sum.c8 = sum(Sum.c8, gVal);
                } else {
                    Sum.d8 = sum(Sum.d8, gVal);
                }
            } else if (eq35kV) {
                if (le20141231) {
                    Sum.c9 = sum(Sum.c9, gVal);
                } else {
                    Sum.d9 = sum(Sum.d9, gVal);
                }
            }

        }
    }

    private static void fillSumSheet() {
        Workbook workbook = mainBook.getXSSFWorkbook(); //直接用SXSSFWorkbook不能获取到值

        Sheet sumSheet2 = workbook.getSheetAt(1);
        Row row = sumSheet2.getRow(5);
        row.getCell(2).setCellValue(Sum.c6.doubleValue());
        row.getCell(3).setCellValue(Sum.d6.doubleValue());

        row = sumSheet2.getRow(6);
        row.getCell(2).setCellValue(Sum.c7.doubleValue());
        row.getCell(3).setCellValue(Sum.d7.doubleValue());

        row = sumSheet2.getRow(7);
        row.getCell(2).setCellValue(Sum.c8.doubleValue());
        row.getCell(3).setCellValue(Sum.d8.doubleValue());

        row = sumSheet2.getRow(8);
        row.getCell(2).setCellValue(Sum.c9.doubleValue());
        row.getCell(3).setCellValue(Sum.d9.doubleValue());
    }

    private static BigDecimal sum(BigDecimal oldVal, String strVal) {
        return oldVal.add(new BigDecimal(strVal));
    }

    private static String copyCell(Workbook newWorkbook, Cell oldCell, Cell newCell, Map<Integer, CellStyle> styleMap) {
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

        String strVal = null;

        CellType oldCellType = oldCell.getCellType();
        if (oldCellType == CellType.FORMULA) {
            newCell.setCellFormula(oldCell.getCellFormula());
            oldCellType = oldCell.getCachedFormulaResultType();
        }
        switch (oldCellType) {
            case STRING:
                RichTextString richStringVal = oldCell.getRichStringCellValue();
                newCell.setCellValue(richStringVal);
                strVal = richStringVal.getString();
                break;
            case NUMERIC:
                double doubleVal = oldCell.getNumericCellValue();
                newCell.setCellValue(doubleVal);
                strVal = Double.toString(doubleVal);
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

        return strVal;
    }

    private static String getCellVal(Cell cell) {
        if (cell == null) return "0";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getRichStringCellValue().getString();
            case NUMERIC:
                return Double.toString(cell.getNumericCellValue());
        }
        return "0";
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
