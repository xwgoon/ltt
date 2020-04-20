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
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
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


//        System.out.println(ChronoUnit.DAYS.between(LocalDate.of(2003, 4, 1), LocalDate.of(1900, 1, 1)));

        System.out.println((int) 'B');
    }

    private static boolean isFirstExcel = true;
    private static int mainSheetLastRowNum;
    private static String[] cellValArr = new String[30];
    private static final Map<String, BigDecimal> valMap = new HashMap<>();

    //资产类别：输电线路,变电设备,配电线路,配电设备-其他,配电设备-电动汽车充换电设备,用电计量设备,通信线路及设备,
    // 自动化控制设备、信息设备及仪器仪表,发电及供热设备,水工机械设备,制造及检修维护设备,生产管理用工器具,运输设备,辅助生产用设备及器具,
    // 房屋,建筑物,土地
    private static boolean eq输电线路;
    private static boolean eq变电设备;
    private static boolean eq配电线路;
    private static boolean eq配电设备其他;
    private static boolean eq配电设备电动汽车充换电设备;
    private static boolean eq用电计量设备;
    private static boolean eq通信线路及设备;
    private static boolean eq自动化控制设备信息设备及仪器仪表;
    private static boolean eq发电及供热设备;
    private static boolean eq水工机械设备;
    private static boolean eq制造及检修维护设备;
    private static boolean eq生产管理用工器具;
    private static boolean eq运输设备;
    private static boolean eq辅助生产用设备及器具;
    private static boolean eq房屋;
    private static boolean eq建筑物;
    private static boolean eq土地;

    //电压等级：500kV,220kV,110kV,35kV,10kV,10kV以下
    private static boolean eq500KV;
    private static boolean eq220KV;
    private static boolean eq110KV;
    private static boolean eq35KV;
    private static boolean eq10KV;
    private static boolean eq10KV以下;

    //资本化日期
    private static boolean le20141231;

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

            completeMapVal();
            fillSheet2();

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
        Arrays.fill(cellValArr, null);
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

        //资产类别
        eq输电线路 = false;
        eq变电设备 = false;
        eq配电线路 = false;
        eq配电设备其他 = false;
        eq配电设备电动汽车充换电设备 = false;
        eq用电计量设备 = false;
        eq通信线路及设备 = false;
        eq自动化控制设备信息设备及仪器仪表 = false;
        eq发电及供热设备 = false;
        eq水工机械设备 = false;
        eq制造及检修维护设备 = false;
        eq生产管理用工器具 = false;
        eq运输设备 = false;
        eq辅助生产用设备及器具 = false;
        eq房屋 = false;
        eq建筑物 = false;
        eq土地 = false;

        switch (cellValArr[1]) {
            case "输电线路":
                eq输电线路 = true;
                break;
            case "变电设备":
                eq变电设备 = true;
                break;
            case "配电线路":
                eq配电线路 = true;
                break;
            case "配电设备-其他":
                eq配电设备其他 = true;
                break;
            case "配电设备-电动汽车充换电设备":
                eq配电设备电动汽车充换电设备 = true;
                break;
            case "用电计量设备":
                eq用电计量设备 = true;
                break;
            case "通信线路及设备":
                eq通信线路及设备 = true;
                break;
            case "自动化控制设备、信息设备及仪器仪表":
                eq自动化控制设备信息设备及仪器仪表 = true;
                break;
            case "发电及供热设备":
                eq发电及供热设备 = true;
                break;
            case "水工机械设备":
                eq水工机械设备 = true;
                break;
            case "制造及检修维护设备":
                eq制造及检修维护设备 = true;
                break;
            case "生产管理用工器具":
                eq生产管理用工器具 = true;
                break;
            case "辅助生产用设备及器具":
                eq辅助生产用设备及器具 = true;
                break;
            case "房屋":
                eq房屋 = true;
                break;
            case "建筑物":
                eq建筑物 = true;
                break;
            case "土地":
                eq土地 = true;
                break;
        }


        //电压等级
        eq500KV = false;
        eq220KV = false;
        eq110KV = false;
        eq35KV = false;
        eq10KV = false;
        eq10KV以下 = false;

        switch (cellValArr[3].toUpperCase()) {
            case "500KV":
                eq500KV = true;
                break;
            case "220KV":
                eq220KV = true;
                break;
            case "110KV":
                eq110KV = true;
                break;
            case "35KV":
                eq35KV = true;
                break;
            case "10KV":
                eq10KV = true;
                break;
            case "10KV以下":
                eq10KV以下 = true;
                break;
        }

        //资本化日期（2014-12-31，poi获取到的值是42004.0）
        le20141231 = "42004.0".compareTo(cellValArr[4]) >= 0;

        calcCol("C", "D", cellValArr[6]);
        calcCol("F", "G", cellValArr[9]);
        calcCol("I", "J", cellValArr[8]);
        calcCol("L", "M", cellValArr[11]);
        calcCol("O", "P", cellValArr[12]);

    }

    private static void calcCol(String le20141231Col, String gt20141231Col, String val) {
        String col = le20141231 ? le20141231Col : gt20141231Col;
        String aVal = cellValArr[0];
        if (eq输电线路) {
            if (eq500KV) {
                calcMap(col, 6, val);
            } else if (eq220KV) {
                calcMap(col, 7, val);
            } else if (eq110KV) {
                calcMap(col, 8, val);
            } else if (eq35KV) {
                calcMap(col, 9, val);
            }
        } else if (eq变电设备) {
            if (eq500KV) {
                calcMap(col, 11, val);
            } else if (eq220KV) {
                calcMap(col, 12, val);
            } else if (eq110KV) {
                calcMap(col, 13, val);
            } else if (eq35KV) {
                calcMap(col, 14, val);
            } else if (eq10KV) {
                calcMap(col, 15, val);
            }
        } else if (eq配电线路) {
            if (eq35KV) {
                calcMap(col, 18, val);
            } else if (eq10KV) {
                calcMap(col, 19, val);
            } else if (eq10KV以下) {
                calcMap(col, 20, val);
            }
        } else if (eq配电设备其他) {
            if (eq35KV) {
                calcMap(col, 22, val);
            } else if (eq10KV) {
                calcMap(col, 23, val);
            } else if (eq10KV以下) {
                calcMap(col, 24, val);
            }
        } else if (eq配电设备电动汽车充换电设备) {
            calcMap(col, 25, val);
        } else if (eq用电计量设备) {
            calcMap(col, 26, val);
        } else if (eq通信线路及设备) {
            calcMap(col, 27, val);
        } else if (eq自动化控制设备信息设备及仪器仪表) {
            if (aVal.startsWith("2001")) {
                calcMap(col, 29, val);
            } else if (aVal.startsWith("2004")) {
                calcMap(col, 30, val);
            } else if (aVal.startsWith("2099")) {
                calcMap(col, 31, val);
            } else if (aVal.startsWith("2003")) {
                calcMap(col, 32, val);
            } else if (aVal.startsWith("2002")) {
                calcMap(col, 33, val);
            }
        } else if (eq发电及供热设备) {
            if (aVal.startsWith("2101")) {
                calcMap(col, 35, val);
            } else if (aVal.startsWith("2102")) {
                calcMap(col, 36, val);
            } else if (aVal.startsWith("2103")) {
                calcMap(col, 37, val);
            } else if (aVal.startsWith("2104")) {
                calcMap(col, 38, val);
            } else if (aVal.startsWith("2105")) {
                calcMap(col, 39, val);
            } else if (aVal.startsWith("2113")) {
                calcMap(col, 40, val);
            } else if (aVal.startsWith("2106")) {
                calcMap(col, 41, val);
            } else if (aVal.startsWith("2107")) {
                calcMap(col, 42, val);
            } else if (aVal.startsWith("2108")) {
                calcMap(col, 43, val);
            } else if (aVal.startsWith("2109")) {
                calcMap(col, 44, val);
            } else if (aVal.startsWith("2110")) {
                calcMap(col, 45, val);
            } else if (aVal.startsWith("2111")) {
                calcMap(col, 46, val);
            } else if (aVal.startsWith("2112")) {
                calcMap(col, 47, val);
            } else if (aVal.startsWith("2199")) {
                calcMap(col, 48, val);
            }
        } else if (eq水工机械设备) {
            calcMap(col, 49, val);
        } else if (eq制造及检修维护设备) {
            calcMap(col, 50, val);
        } else if (eq生产管理用工器具) {
            calcMap(col, 51, val);
        } else if (eq运输设备) {
            if (aVal.startsWith("2501")) {
                calcMap(col, 53, val);
            } else if (aVal.startsWith("2502")) {
                calcMap(col, 54, val);
            } else if (aVal.startsWith("2503")) {
                calcMap(col, 55, val);
            } else if (aVal.startsWith("2504")) {
                calcMap(col, 56, val);
            } else if (aVal.startsWith("2599")) {
                calcMap(col, 57, val);
            }
        } else if (eq辅助生产用设备及器具) {
            calcMap(col, 58, val);
        } else if (eq房屋) {
            calcMap(col, 59, val);
        } else if (eq建筑物) {
            calcMap(col, 60, val);
        } else if (eq土地) {
            calcMap(col, 61, val);
        }
    }

    private static boolean isBlank(String str) {
        int strLen;
        if (str != null && (strLen = str.length()) != 0) {
            for (int i = 0; i < strLen; ++i) {
                if (!Character.isWhitespace(str.charAt(i))) {
                    return false;
                }
            }

            return true;
        } else {
            return true;
        }
    }

    private static void calcMap(String col, int row, String val) {
        BigDecimal decimalVal = isBlank(val) ? BigDecimal.ZERO : new BigDecimal(val);
        valMap.merge(col + row, decimalVal, BigDecimal::add);
    }

    private static void calcMap(String col, int row, BigDecimal val) {
        valMap.merge(col + row, val, BigDecimal::add);
    }

    private static void subColVal(String resultCol, String col1, String col2, int row) {
        BigDecimal col1Val = valMap.getOrDefault(col1 + row, BigDecimal.ZERO);
        BigDecimal col2Val = valMap.getOrDefault(col2 + row, BigDecimal.ZERO);
        BigDecimal resultVal = col1Val.subtract(col2Val);
        calcMap(resultCol, row, resultVal);
    }

    private static void completeMapVal() {

        //三、配电线路及设备 - 2.配电设备 - 10千伏以下，需要加上 配电设备-电动汽车充换电设备 项
        String[] cols = {"C", "D", "F", "G", "I", "J", "L", "M", "O", "P"};
        for (String col : cols) {
            sumColVal(col, 24, 24, 25);
        }

        //年初净值、年末净值
        for (int i = 6; i <= 61; i++) {
            subColVal("R", "C", "I", i);
            subColVal("S", "D", "J", i);
            subColVal("U", "F", "L", i);
            subColVal("V", "G", "M", i);
        }

        cols = new String[]{"C", "D", "F", "G", "I", "J", "L", "M", "O", "P", "R", "S", "U", "V"};
        int[][] rowStartEndArr = {
                {5, 9}, //一、输电线路
                {10, 15}, //二、变电设备
                {17, 20}, //1.配电线路
                {21, 24}, //2.配电设备
                {28, 33}, //六、自动化控制设备、信息设备及仪器仪表
                {34, 48}, //七、发电及供热设备
                {52, 57}, //十一、运输设备
        };

        for (int[] startEnd : rowStartEndArr) {
            for (String col : cols) {
                sumColVal(col, startEnd[0], startEnd[1]);
            }
        }

        //三、配电线路及设备
        int[] heJiRows = {17, 21};
        for (String col : cols) {
            sumColVal(col, 16, heJiRows);
        }

        //合计
        heJiRows = new int[]{5, 10, 16, 26, 27, 28, 34, 49, 50, 51, 52, 58, 59, 60, 61};
        for (String col : cols) {
            sumColVal(col, 62, heJiRows);
        }

        //专用设备合计
        heJiRows = new int[]{5, 10, 16, 30, 31, 32, 37, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 49, 53, 54, 56};
        for (String col : cols) {
            sumColVal(col, 63, heJiRows);
        }

        //通用设备合计
        heJiRows = new int[]{26, 27, 29, 33, 35, 36, 44, 50, 51, 55, 57, 58};
        for (String col : cols) {
            sumColVal(col, 64, heJiRows);
        }

        //通用设备合计
        heJiRows = new int[]{59, 60, 61};
        for (String col : cols) {
            sumColVal(col, 65, heJiRows);
        }

        //计算列合计值
        for (int i = 5; i <= 65; i++) {
            for (char c = 'B'; c <= 'T'; c += 3) {
                sumRowVal(i, c, c + 2);
            }
        }
    }

    private static void fillSheet2() {
        Workbook workbook = mainBook.getXSSFWorkbook(); //直接用SXSSFWorkbook不能获取到值
        Sheet sheet2 = workbook.getSheetAt(1);

        Row row;
        Cell cell;
        String position;
        for (int i = 5; i <= 65; i++) {
            row = sheet2.getRow(i - 1);
            for (int j = 1; j <= 21; j++) {
                position = (char) ('A' + j) + String.valueOf(i);
                BigDecimal bigDecimalVal = valMap.get(position);
                double doubleVal = bigDecimalVal == null ? 0 : bigDecimalVal.doubleValue();
                cell = row.getCell(j);
                cell.setCellValue(doubleVal);
            }
        }
    }

    private static void sumColVal(String col, int rowStart, int rowEnd) {
        BigDecimal sumVal = BigDecimal.ZERO;
        for (int i = rowStart + 1; i <= rowEnd; i++) {
            BigDecimal val = valMap.get(col + i);
            if (val != null) {
                sumVal = sumVal.add(val);
            }
        }
        valMap.put(col + rowStart, sumVal);
    }

    private static void sumColVal(String col, int resultRow, int... sumRows) {
        BigDecimal sumVal = BigDecimal.ZERO;
        for (int row : sumRows) {
            BigDecimal val = valMap.get(col + row);
            if (val != null) {
                sumVal = sumVal.add(val);
            }
        }
        valMap.put(col + resultRow, sumVal);
    }

    private static void sumRowVal(int row, int colStart, int colEnd) {
        String rowStr = String.valueOf(row);
        BigDecimal sumVal = BigDecimal.ZERO;
        for (int i = colStart + 1; i <= colEnd; i++) {
            BigDecimal val = valMap.get((char) i + rowStr);
            if (val != null) {
                sumVal = sumVal.add(val);
            }
        }
        valMap.put((char) colStart + rowStr, sumVal);
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
