package com.xw.ltt;

import com.xw.ltt.excel.ExcelUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

public class Test {

    public static void main(String[] args) {

        List<InputStream> excelFiles = new ArrayList<>();

        Path excelDir = Paths.get("原始数据");
        try (Stream<Path> filePaths = Files.list(excelDir)) {
            filePaths.forEach(filePath -> {
                try {
                    InputStream inputStream = Files.newInputStream(filePath);
                    excelFiles.add(inputStream);
                } catch (IOException e) {
                    e.printStackTrace();
                    throw new RuntimeException("读取文件出错");
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("遍历文件出错");
        }

        File file = new File("运行结果/合并结果.xlsx");
        try {
            ExcelUtil.mergeExcelFiles(file, excelFiles);
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException("合并Excel出错");
        }
    }

}
