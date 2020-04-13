package com.xw.ltt;

import com.xw.ltt.excel.ExcelUtil;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Stream;

public class Test {

    public static final String WORK_DIR;

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
//        IntHolder time = new IntHolder(1);
//        Timer timer = new Timer(true);
//        timer.schedule(new TimerTask() {
//            @Override
//            public void run() {
//                System.out.print(time.value++ + "s>");
//            }
//        }, 1000, 1000);
        System.out.println("开始计算...");
        long startTime = System.currentTimeMillis();

        Map<String, InputStream> excelFiles = new LinkedHashMap<>();
        Path excelDir = Paths.get(WORK_DIR + "原始数据");
        try (Stream<Path> filePaths = Files.list(excelDir)) {
            filePaths.forEachOrdered(filePath -> {
                try {
                    InputStream inputStream = Files.newInputStream(filePath);
                    excelFiles.put(filePath.getFileName().toString(), inputStream);
                } catch (IOException e) {
                    e.printStackTrace();
                    throw new RuntimeException("读取文件出错");
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("遍历文件出错");
        }

        File file = new File(WORK_DIR + "计算结果/合并结果.xlsx");
        try {
            ExcelUtil.mergeExcelFiles(file, excelFiles);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("合并Excel出错");
        }

        System.out.println("计算完成，共耗时" + (System.currentTimeMillis() - startTime) / 1000 + "秒。\n");

        System.out.println("请按回车键结束...");
        System.in.read();
    }

}
