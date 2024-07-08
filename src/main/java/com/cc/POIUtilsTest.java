package com.cc;

import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;


public class POIUtilsTest {


    @Test
    public void test() throws IOException {

        // Excel文件
        String path = "E:\\temp\\spark\\西安大数据.xlsx";

        // 加载Excel
        Sheet sheet = POIUtils.excelInit(path, "sheet1", true);

        // 获取行列数
        System.out.println(POIUtils.getRowsNum(sheet));
        System.out.println(POIUtils.getColsNum(sheet));

        // 读取Excel
        String rows = POIUtils.excelReader(sheet);
        System.out.println(rows);
        List<String> lines = Arrays.asList(rows.split("\\n"));
//        List<List<String>> lines = POIUtils.excelParser(sheet);
//        lines.forEach(System.out::println);

        System.out.println("总行数（不包括空行）：" + lines.size());

        // 释放资源
        POIUtils.excelDestroy();

        // 写入CSV/文本文件
//        POIUtils.excelWriter(rows, "C:\\Users\\cc\\Desktop\\out.csv", true);

    }


}

