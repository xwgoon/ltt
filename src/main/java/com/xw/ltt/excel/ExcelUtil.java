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

//        System.out.println((int) 'A');
//        int x=25; //0<x<
//        String base26Str=Integer.toString(x,26);
////        String base26Str1=Integer.toString(x,26);
//        System.out.println(base26Str);
////        char[] chars=str.toCharArray();
//        if(x<26){
//            System.out.println((char)('A'+x));
//        }else{
//            String[] base26Arr=base26Str.split("");
//            System.out.println(base26Arr);
//            for (String str : base26Arr) {
//                System.out.println(Integer.valueOf(str));
//                char letter=(char)(Integer.valueOf(str)+64);
//                System.out.println(letter);
//            }
//        }


//        int[] arr=new int[10];
//        Arrays.fill(arr,-1);
//        int idx=9;
//        int mod=26;
//
//        int x=26;
//
//        while (x>0){
//            arr[idx--]=x%mod;
//            x=x/mod;
//        }
////        arr[idx]=x;
//
//        System.out.println(Arrays.toString(arr));

//        char[] chars=new char[10];
//        for (int i = arr.length - 1; i >= 0; i--) {
//            int val=arr[i];
//            if(val==-1) break;
//            if(i<arr.length - 1){
//                val-=1;
//            }
//            chars[i]=(char) (val+'A');
//        }
//
//        System.out.println(Arrays.toString(chars));


//        StringBuilder columnName = new StringBuilder();
//        int dividend = 101;
//        int modulo;
//        while (dividend > 0) {
//            modulo = (dividend - 1) % 26;
//            columnName.insert(0, (char) (65 + modulo));
//            dividend = (dividend - modulo) / 26;
//        }
//        System.out.println(columnName);

    }

    private static boolean isFirstExcel = true;
    private static int mainSheetLastRowNum;
    private static String[] cellValArr = new String[30];
    private static final Map<String, BigDecimal> S1ValMap = new HashMap<>();
    private static final Map<String, BigDecimal> S0ValMap = new HashMap<>();

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
            Path templatePath = Paths.get(Test.WORK_DIR + "模板/卡片数据模板.xlsx");
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

            if (Test.isCard) {
                completeS1ValMap();
                completeS0ValMap();

                fillSheet(1, 65, 21);
                fillSheet(0, 37, 41);
            }

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
        calcCell();
    }

    private static void calcCell() {

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

        calcS1Cell();

        calcS0Cell();

    }

    private static void calcS1Cell() {
        calcS1Col("C", "D", cellValArr[6]);
        calcS1Col("F", "G", cellValArr[9]);
        calcS1Col("I", "J", cellValArr[8]);
        calcS1Col("L", "M", cellValArr[11]);
        calcS1Col("O", "P", cellValArr[12]);
    }

    private static void calcS0Cell() {
        //年初-其中：已提足折旧资产原值
        calcS0Col("C", cellValArr[7]);

        //年初-累计折旧
        calcS0Col("V", cellValArr[8]);

        //本年增加-其中：本年计提
        calcS0Col("Y", cellValArr[12]);

        //年末-累计折旧
        calcS0Col("AL", cellValArr[11]);

        String col1;
        String col2;

        //本年增加
        if (!isBlank(cellValArr[13])) {
            col1 = null; //资产原值
            col2 = null; //累计折旧
            switch (cellValArr[13]) { //资产变动类型
                case "工程项目":
                case "零购项目":
                case "用户资产":
                case "盘盈资产":
                case "非货币性交易":
                case "捐赠增加":
                    col1 = "E";
                    col2 = "X";
                    break;
                case "子分公司间划转资产":
                    col1 = "F";
                    col2 = "Z";
                    break;
                case "省外划拨":
                    col1 = "G";
                    col2 = "AA";
                    break;
                case "省内地市间划拨":
                    col1 = "H";
                    col2 = "AB";
                    break;
                case "拆分合并重分类":
                case "地市内划拨":
                    col1 = "I";
                    col2 = "AC";
                    break;
            }
            calcS0Col(col1, cellValArr[14]);
            calcS0Col(col2, cellValArr[15]);
        }

        //本年减少
        if (!isBlank(cellValArr[16])) {
            col1 = null; //资产原值
            col2 = null; //累计折旧
            switch (cellValArr[16]) { //资产变动类型
                case "报废":
                    col1 = "K";
                    col2 = "AE";
                    break;
                case "出售":
                    col1 = "L";
                    col2 = "AF";
                    break;
                case "三供一业无偿划出":
                    col1 = "M";
                    col2 = "AG";
                    break;
                case "子分公司间划转资产":
                    col1 = "N";
                    col2 = "AH";
                    break;
                case "省外划拨":
                    col1 = "O";
                    col2 = "AI";
                    break;
                case "省内地市间划拨":
                    col1 = "P";
                    col2 = "AJ";
                    break;
                case "拆分合并重分类":
                case "地市内划拨":
                    col1 = "Q";
                    col2 = "AK";
                    break;
            }
            calcS0Col(col1, cellValArr[17]);
            calcS0Col(col2, cellValArr[18]);
        }

        //年末-其中：已提足折旧资产原值
        calcS0Col("S", cellValArr[10]);

        //逾龄资产预计	-预计2020年末逾龄资产
        calcS0Col("AO", cellValArr[22]);

        //逾龄资产预计	-预计2021年末逾龄资产
        calcS0Col("AP", cellValArr[23]);
    }

    private static void calcS0Col(String col, String val) {
        if (eq输电线路) {
            if (eq500KV) {
                calcS0Map(col, 6, val);
            } else if (eq220KV) {
                calcS0Map(col, 7, val);
            } else if (eq110KV) {
                calcS0Map(col, 8, val);
            } else if (eq35KV) {
                calcS0Map(col, 9, val);
            }
        } else if (eq变电设备) {
            if (eq500KV) {
                calcS0Map(col, 11, val);
            } else if (eq220KV) {
                calcS0Map(col, 12, val);
            } else if (eq110KV) {
                calcS0Map(col, 13, val);
            } else if (eq35KV) {
                calcS0Map(col, 14, val);
            } else if (eq10KV) {
                calcS0Map(col, 15, val);
            }
        } else if (eq配电线路) {
            if (eq35KV) {
                calcS0Map(col, 17, val);
            } else if (eq10KV) {
                calcS0Map(col, 18, val);
            } else if (eq10KV以下) {
                calcS0Map(col, 19, val);
            }
        } else if (eq配电设备其他) {
            if (eq35KV) {
                calcS0Map(col, 21, val);
            } else if (eq10KV) {
                calcS0Map(col, 22, val);
            } else if (eq10KV以下) {
                calcS0Map(col, 23, val);
            }
        } else if (eq配电设备电动汽车充换电设备) {
            calcS0Map(col, 24, val);
        } else if (eq用电计量设备) {
            calcS0Map(col, 25, val);
        } else if (eq通信线路及设备) {
            calcS0Map(col, 26, val);
        } else if (eq自动化控制设备信息设备及仪器仪表) {
            calcS0Map(col, 27, val);
        } else if (eq发电及供热设备) {
            calcS0Map(col, 28, val);
        } else if (eq水工机械设备) {
            calcS0Map(col, 29, val);
        } else if (eq制造及检修维护设备) {
            calcS0Map(col, 30, val);
        } else if (eq生产管理用工器具) {
            calcS0Map(col, 31, val);
        } else if (eq运输设备) {
            calcS0Map(col, 32, val);
        } else if (eq辅助生产用设备及器具) {
            calcS0Map(col, 33, val);
        } else if (eq房屋) {
            calcS0Map(col, 34, val);
        } else if (eq建筑物) {
            calcS0Map(col, 35, val);
        } else if (eq土地) {
            calcS0Map(col, 36, val);
        }
    }

    private static void calcS1Col(String le20141231Col, String gt20141231Col, String val) {
        String col = le20141231 ? le20141231Col : gt20141231Col;
        String aVal = cellValArr[0];
        if (eq输电线路) {
            if (eq500KV) {
                calcS1Map(col, 6, val);
            } else if (eq220KV) {
                calcS1Map(col, 7, val);
            } else if (eq110KV) {
                calcS1Map(col, 8, val);
            } else if (eq35KV) {
                calcS1Map(col, 9, val);
            }
        } else if (eq变电设备) {
            if (eq500KV) {
                calcS1Map(col, 11, val);
            } else if (eq220KV) {
                calcS1Map(col, 12, val);
            } else if (eq110KV) {
                calcS1Map(col, 13, val);
            } else if (eq35KV) {
                calcS1Map(col, 14, val);
            } else if (eq10KV) {
                calcS1Map(col, 15, val);
            }
        } else if (eq配电线路) {
            if (eq35KV) {
                calcS1Map(col, 18, val);
            } else if (eq10KV) {
                calcS1Map(col, 19, val);
            } else if (eq10KV以下) {
                calcS1Map(col, 20, val);
            }
        } else if (eq配电设备其他) {
            if (eq35KV) {
                calcS1Map(col, 22, val);
            } else if (eq10KV) {
                calcS1Map(col, 23, val);
            } else if (eq10KV以下) {
                calcS1Map(col, 24, val);
            }
        } else if (eq配电设备电动汽车充换电设备) {
            calcS1Map(col, 25, val);
        } else if (eq用电计量设备) {
            calcS1Map(col, 26, val);
        } else if (eq通信线路及设备) {
            calcS1Map(col, 27, val);
        } else if (eq自动化控制设备信息设备及仪器仪表) {
            if (aVal.startsWith("2001")) {
                calcS1Map(col, 29, val);
            } else if (aVal.startsWith("2004")) {
                calcS1Map(col, 30, val);
            } else if (aVal.startsWith("2099")) {
                calcS1Map(col, 31, val);
            } else if (aVal.startsWith("2003")) {
                calcS1Map(col, 32, val);
            } else if (aVal.startsWith("2002")) {
                calcS1Map(col, 33, val);
            }
        } else if (eq发电及供热设备) {
            if (aVal.startsWith("2101")) {
                calcS1Map(col, 35, val);
            } else if (aVal.startsWith("2102")) {
                calcS1Map(col, 36, val);
            } else if (aVal.startsWith("2103")) {
                calcS1Map(col, 37, val);
            } else if (aVal.startsWith("2104")) {
                calcS1Map(col, 38, val);
            } else if (aVal.startsWith("2105")) {
                calcS1Map(col, 39, val);
            } else if (aVal.startsWith("2113")) {
                calcS1Map(col, 40, val);
            } else if (aVal.startsWith("2106")) {
                calcS1Map(col, 41, val);
            } else if (aVal.startsWith("2107")) {
                calcS1Map(col, 42, val);
            } else if (aVal.startsWith("2108")) {
                calcS1Map(col, 43, val);
            } else if (aVal.startsWith("2109")) {
                calcS1Map(col, 44, val);
            } else if (aVal.startsWith("2110")) {
                calcS1Map(col, 45, val);
            } else if (aVal.startsWith("2111")) {
                calcS1Map(col, 46, val);
            } else if (aVal.startsWith("2112")) {
                calcS1Map(col, 47, val);
            } else if (aVal.startsWith("2199")) {
                calcS1Map(col, 48, val);
            }
        } else if (eq水工机械设备) {
            calcS1Map(col, 49, val);
        } else if (eq制造及检修维护设备) {
            calcS1Map(col, 50, val);
        } else if (eq生产管理用工器具) {
            calcS1Map(col, 51, val);
        } else if (eq运输设备) {
            if (aVal.startsWith("2501")) {
                calcS1Map(col, 53, val);
            } else if (aVal.startsWith("2502")) {
                calcS1Map(col, 54, val);
            } else if (aVal.startsWith("2503")) {
                calcS1Map(col, 55, val);
            } else if (aVal.startsWith("2504")) {
                calcS1Map(col, 56, val);
            } else if (aVal.startsWith("2599")) {
                calcS1Map(col, 57, val);
            }
        } else if (eq辅助生产用设备及器具) {
            calcS1Map(col, 58, val);
        } else if (eq房屋) {
            calcS1Map(col, 59, val);
        } else if (eq建筑物) {
            calcS1Map(col, 60, val);
        } else if (eq土地) {
            calcS1Map(col, 61, val);
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
        }
        return true;
    }

    private static void calcS1Map(String col, int row, String val) {
        BigDecimal decimalVal = isBlank(val) ? BigDecimal.ZERO : new BigDecimal(val);
        calcMap(S1ValMap, col, row, decimalVal);
    }

    private static void calcS0Map(String col, int row, String val) {
        BigDecimal decimalVal = isBlank(val) ? BigDecimal.ZERO : new BigDecimal(val);
        calcMap(S0ValMap, col, row, decimalVal);
    }

    private static void calcMap(Map<String, BigDecimal> map, String col, int row, BigDecimal val) {
        map.merge(col + row, val, BigDecimal::add);
    }

    private static void subColVal(Map<String, BigDecimal> map, String resultCol, String col1, String col2, int row) {
        BigDecimal col1Val = map.getOrDefault(col1 + row, BigDecimal.ZERO);
        BigDecimal col2Val = map.getOrDefault(col2 + row, BigDecimal.ZERO);
        BigDecimal resultVal = col1Val.subtract(col2Val);
        calcMap(map, resultCol, row, resultVal);
    }

    private static void S0PeriodCheck(String resultCol, String addCol1, String addCol2,
                                      String subCol1, String subCol2, int row) {
        BigDecimal addCol1Val = S0ValMap.getOrDefault(addCol1 + row, BigDecimal.ZERO);
        BigDecimal addCol2Val = S0ValMap.getOrDefault(addCol2 + row, BigDecimal.ZERO);
        BigDecimal subCol1Val = S0ValMap.getOrDefault(subCol1 + row, BigDecimal.ZERO);
        BigDecimal subCol2Val = S0ValMap.getOrDefault(subCol2 + row, BigDecimal.ZERO);
        BigDecimal resultVal = addCol1Val.add(addCol2Val).subtract(subCol1Val).subtract(subCol2Val);
        calcMap(S0ValMap, resultCol, row, resultVal);

    }

    private static void completeS1ValMap() {

        //三、配电线路及设备 - 2.配电设备 - 10千伏以下，需要加上 配电设备-电动汽车充换电设备 项
        String[] cols = {"C", "D", "F", "G", "I", "J", "L", "M", "O", "P"};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 24, 24, 25);
        }

        //年初净值、年末净值
        for (int i = 6; i <= 61; i++) {
            subColVal(S1ValMap, "R", "C", "I", i);
            subColVal(S1ValMap, "S", "D", "J", i);
            subColVal(S1ValMap, "U", "F", "L", i);
            subColVal(S1ValMap, "V", "G", "M", i);
        }

        cols = new String[]{"C", "D", "F", "G", "I", "J", "L", "M", "O", "P", "R", "S", "U", "V"};
        int[][] rowStartEnds = {
                {5, 9}, //一、输电线路
                {10, 15}, //二、变电设备
                {17, 20}, //1.配电线路
                {21, 24}, //2.配电设备
                {28, 33}, //六、自动化控制设备、信息设备及仪器仪表
                {34, 48}, //七、发电及供热设备
                {52, 57}, //十一、运输设备
        };
        sumColVal(S1ValMap, cols, rowStartEnds);

        //三、配电线路及设备
        int[] heJiRows = {17, 21};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 16, heJiRows);
        }

        //合计
        heJiRows = new int[]{5, 10, 16, 26, 27, 28, 34, 49, 50, 51, 52, 58, 59, 60, 61};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 62, heJiRows);
        }

        //专用设备合计
        heJiRows = new int[]{5, 10, 16, 30, 31, 32, 37, 38, 39, 40, 41, 42, 43, 45, 46, 47, 48, 49, 53, 54, 56};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 63, heJiRows);
        }

        //通用设备合计
        heJiRows = new int[]{26, 27, 29, 33, 35, 36, 44, 50, 51, 55, 57, 58};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 64, heJiRows);
        }

        //通用设备合计
        heJiRows = new int[]{59, 60, 61};
        for (String col : cols) {
            sumColVal(S1ValMap, col, 65, heJiRows);
        }

        //计算列合计值
        for (int i = 5; i <= 65; i++) {
            for (int c = 1; c <= 19; c += 3) {
                sumRowVal(S1ValMap, i, c, c + 2);
            }
        }
    }

    private static void completeS0ValMap() {
        String S0S1ColMap[][] = {
                {"B", "B"}, //年初-资产原值
                {"R", "E"}, //年末-资产原值
        };
        int S0S1RowMap[][] = {
                {5, 5}, {6, 6}, {7, 7}, {8, 8}, {9, 9}, {10, 10}, {11, 11}, {12, 12}, {13, 13}, {14, 14}, {15, 15},
                {16, 17}, {17, 18}, {18, 19}, {19, 20}, {20, 21}, {21, 22}, {22, 23}, {23, 24}, {24, 25}, {25, 26},
                {26, 27}, {27, 28}, {28, 34}, {29, 49}, {30, 50}, {31, 51}, {32, 52}, {33, 58}, {34, 59}, {35, 60},
                {36, 61}, {37, 62}
        };
        for (String[] colMap : S0S1ColMap) {
            for (int[] rowMap : S0S1RowMap) {
                copyValFromS1ToS0(colMap[1], rowMap[1], colMap[0], rowMap[0]);
            }
        }

        //四、配电设备-10千伏以下，需要加上【配电设备-电动汽车充换电设备】项
        String[] cols = new String[]{"C", "E", "F", "G", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "S", "V", "X",
                "Y", "Z", "AA", "AB", "AC", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AO", "AP"};
        for (String col : cols) {
            sumColVal(S0ValMap, col, 23, 23, 24);
        }

        cols = new String[50];
        int k = 0;
        for (int i = 2; i <= 41; i++) {
            if (i != 3 && i != 9 && i != 17 && i != 19 && i != 20 && i != 22 && i != 29 && i != 38 && i != 39) {
                cols[k++] = getColName(i);
            }
        }
        int[][] rowStartEnds = {
                {5, 9}, //一、输电线路
                {10, 15}, //二、变电设备
                {16, 19}, //三、配电线路
                {20, 23}, //四、配电设备
        };
        sumColVal(S0ValMap, cols, rowStartEnds);

        //计算列合计值
        int[][] colStartEnds = {
                {3, 8}, //本年增加-资产原值合计
                {9, 16}, //本年减少-资产原值合计
                {22, 28}, //本年增加-累计折旧合计
                {29, 36}, //本年减少-累计折旧合计
        };
        for (int i = 5; i <= 36; i++) {
            for (int[] colStartEnd : colStartEnds) {
                sumRowVal(S0ValMap, i, colStartEnd[0], colStartEnd[1]);
            }
        }

        //合计
        int[] heJiRows = new int[]{5, 10, 16, 20, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36};
        for (int i = 2; i <= 41; i++) {
            if (i != 17 && i != 19 && i != 20 && i != 38 && i != 39) {
                sumColVal(S0ValMap, getColName(i), 37, heJiRows);
            }
        }

        //期间数据核对 
        for (int i = 5; i <= 37; i++) {
            S0PeriodCheck("T", "B", "D", "J", "R", i); //资产原值
            S0PeriodCheck("AM", "V", "W", "AD", "AL", i); //累计折旧
            subColVal(S0ValMap, "U", "I", "Q", i); //资产原值
            subColVal(S0ValMap, "AN", "AC", "AK", i); //累计折旧
        }
    }

    private static void copyValFromS1ToS0(String S1Col, int S1Row, String S0Col, int S0Row) {
        BigDecimal S1Val = S1ValMap.get(S1Col + S1Row);
        S0ValMap.put(S0Col + S0Row, S1Val);
    }

    private static void fillSheet(int sheetNum, int rowEnd, int colEnd) {
        Workbook workbook = mainBook.getXSSFWorkbook(); //直接用SXSSFWorkbook不能获取到值
        Sheet sheet = workbook.getSheetAt(sheetNum);
        Map<String, BigDecimal> map = sheetNum == 0 ? S0ValMap : S1ValMap;

        Row row;
        Cell cell;
        String position;
        for (int i = 5; i <= rowEnd; i++) {
            row = sheet.getRow(i - 1);
            for (int j = 1; j <= colEnd; j++) {
                position = getColName(j) + i;
                BigDecimal bigDecimalVal = map.get(position);
                double doubleVal = bigDecimalVal == null ? 0 : bigDecimalVal.doubleValue();
                cell = row.getCell(j);
                cell.setCellValue(doubleVal);
            }
        }
    }

//    private static void sumS1ColVal(String col, int rowStart, int rowEnd) {
//        BigDecimal sumVal = BigDecimal.ZERO;
//        for (int i = rowStart + 1; i <= rowEnd; i++) {
//            BigDecimal val = S1ValMap.get(col + i);
//            if (val != null) {
//                sumVal = sumVal.add(val);
//            }
//        }
//        S1ValMap.put(col + rowStart, sumVal);
//    }

    private static void sumColVal(Map<String, BigDecimal> valMap, String[] cols, int[][] rowStartEnds) {
        for (String col : cols) {
            for (int[] startEnd : rowStartEnds) {
                BigDecimal sumVal = BigDecimal.ZERO;
                int rowStart = startEnd[0];
                int rowEnd = startEnd[1];
                for (int i = rowStart + 1; i <= rowEnd; i++) {
                    BigDecimal val = valMap.get(col + i);
                    if (val != null) {
                        sumVal = sumVal.add(val);
                    }
                }
                valMap.put(col + rowStart, sumVal);
            }
        }
    }

    private static void sumColVal(Map<String, BigDecimal> valMap, String col, int resultRow, int... sumRows) {
        BigDecimal sumVal = BigDecimal.ZERO;
        for (int row : sumRows) {
            BigDecimal val = valMap.get(col + row);
            if (val != null) {
                sumVal = sumVal.add(val);
            }
        }
        valMap.put(col + resultRow, sumVal);
    }

    private static void sumRowVal(Map<String, BigDecimal> valMap, int row, int colStart, int colEnd) {
        BigDecimal sumVal = BigDecimal.ZERO;
        for (int i = colStart + 1; i <= colEnd; i++) {
            BigDecimal val = valMap.get(getColName(i) + row);
            if (val != null) {
                sumVal = sumVal.add(val);
            }
        }
        valMap.put(getColName(colStart) + row, sumVal);
    }

    private static String getColName(int colNum) {
        StringBuilder colName = new StringBuilder();
        int dividend = colNum + 1;
        int modulo;
        while (dividend > 0) {
            modulo = (dividend - 1) % 26;
            colName.insert(0, (char) (65 + modulo));
            dividend = (dividend - modulo) / 26;
        }
        return colName.toString();
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
