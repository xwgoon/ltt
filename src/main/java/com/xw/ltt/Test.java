package com.xw.ltt;

import com.xw.ltt.excel.ExcelUtil;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.InputMismatchException;
import java.util.List;
import java.util.Scanner;
import java.util.stream.Collectors;

public class Test {

    public static final String WORK_DIR;
    public static int sheetNum;
    public static int sheetIndex;
    public static int titleRowNum;
    public static boolean isCard;
    public static boolean isSuccess = true;

    static {
        File jarFile;
        try {
            jarFile = new File(Test.class.getProtectionDomain().getCodeSource().getLocation().toURI());
        } catch (URISyntaxException e) {
            e.printStackTrace();
            throw new RuntimeException("获取jar文件出错");
        }
        WORK_DIR = jarFile.getParent() + "/";
    }

    public static void main(String[] args) throws Exception {
//        List<Path> templatePaths = new ArrayList<>();
//        Path templateDir = Paths.get(WORK_DIR + "表头模板");
//        try (Stream<Path> filePaths = Files.list(templateDir)) {
//            filePaths.forEachOrdered(templatePaths::add);
//        } catch (IOException e) {
//            e.printStackTrace();
//            throw new RuntimeException("遍历模板出错");
//        }

//        for (; ; ) {
//            System.out.println("表头模板：");
//            for (int i = 0; i < templatePaths.size(); i++) {
//                System.out.println(i + 1 + "、" + templatePaths.get(i).getFileName());
//            }
//
//            System.out.print("请输入表头模板的序号（1-" + templatePaths.size() + "的整数）：");
//            Scanner in = new Scanner(System.in);
//            try {
//                int templateNum = in.nextInt();
//                if (templateNum < 1 || templateNum > templatePaths.size()) {
//                    System.out.println();
//                    continue;
//                }
//                templatePath = templatePaths.get(templateNum - 1);
//                break;
//            } catch (InputMismatchException ignored) {
//                System.out.println();
//            }
//        }

        isCard = args.length > 0;

        for (; ; ) {
            System.out.print("请输入表的位置（大于0的整数）：");
            Scanner in = new Scanner(System.in);
            try {
                sheetNum = in.nextInt();
                if (sheetNum <= 0) continue;
                sheetIndex = sheetNum - 1;
                break;
            } catch (InputMismatchException ignored) {
            }
        }

        for (; ; ) {
            System.out.print("请输入表头行数（大于等于0的整数）：");
            Scanner in = new Scanner(System.in);
            try {
                titleRowNum = in.nextInt();
                if (titleRowNum < 0) continue;
                break;
            } catch (InputMismatchException ignored) {
            }
        }

        System.out.println("\n开始计算...");
        long startTime = System.currentTimeMillis();


//        Map<String, InputStream> excelFiles = new LinkedHashMap<>();
//        try (Stream<Path> filePaths = Files.list(excelDir)) {
//            filePaths.forEachOrdered(filePath -> {
//                try {
//                    InputStream inputStream = Files.newInputStream(filePath);
//                    excelFiles.put(filePath.getFileName().toString(), inputStream);
//                } catch (IOException e) {
//                    e.printStackTrace();
//                    throw new RuntimeException("读取文件出错");
//                }
//            });
//        } catch (IOException e) {
//            e.printStackTrace();
//            throw new RuntimeException("遍历文件出错");
//        }


        try {
            Path excelDir = Paths.get(WORK_DIR + "原始数据");
            List<Path> excelPaths = Files.list(excelDir).collect(Collectors.toList());
            File file = new File(WORK_DIR + "计算结果/计算结果.xlsx");
            ExcelUtil.mergeExcelFiles(file, excelPaths);
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (isSuccess) {
            System.out.println("计算成功，用时" + (System.currentTimeMillis() - startTime) / 1000 + "秒。");
        } else {
            System.out.println("计算失败。");
        }

        System.out.println("\n请按回车键结束...");
        System.in.read();
    }

}
