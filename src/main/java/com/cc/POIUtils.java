package com.cc;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;


public class POIUtils {

    private static Workbook workbook = null;
    private static InputStream excel = null;

    // 获取列数（包括空列）
    public static Integer getColsNum(Sheet sheet) {
        // 不合法判断
        if (sheet == null){
            return null;
        }
        return sheet.getLastRowNum() == -1 ? -1 : sheet.getRow(0).getPhysicalNumberOfCells();
    }

    // 获取行数（包括空行）
    public static Integer getRowsNum(Sheet sheet) {
        // 不合法判断
        if (sheet == null){
            return null;
        }
        return sheet.getLastRowNum() == -1 ? -1 : sheet.getPhysicalNumberOfRows();
    }

    // 初始化Excel
    public static Sheet excelInit(String path, String name, Boolean ooxml) throws IOException {
        // 合法性判断
        String fileName = new File(path).getName();
        if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
            return null;
        }
        excel = Files.newInputStream(Paths.get(path));
        // Excel加载
        if (ooxml){
            // Microsoft Office 2007起（xlsx）
            workbook = new XSSFWorkbook(excel);
        } else {
            // Microsoft Office 2007之前（xls）
            workbook = new HSSFWorkbook(excel);
        }
        return workbook.getSheet(name);
    }

    // 读取Excel，返回字符串类型（去除空行）
    public static String excelReader(Sheet sheet) {
        // 不合法判断
        if (sheet == null){
            return null;
        }
        // 行数判断
        if (sheet.getLastRowNum() == -1){
            return "";
        }
        // 遍历行和单元格（除了迭代器和如下遍历，其他遍历可能解析报错）
        StringBuilder builder = new StringBuilder();
        for (Row row : sheet) {
            StringBuilder line = new StringBuilder();
            for (Cell cell : row) {
                line.append(cell.toString().replaceAll("\\n", "\002").trim()).append("\001");
            }
            if (line.toString().trim().length() > 0){
                builder.append(line.toString().trim()).append("\n");
            }
        }
        return builder.toString().trim();
    }

    // 读取Excel，返回列表类型（去除空行）
    public static List<List<String>> excelParser(Sheet sheet) {
        // 不合法判断
        if (sheet == null){
            return null;
        }
        // 行数判断
        if (sheet.getLastRowNum() == -1){
            return null;
        }
        // 遍历行和单元格（除了迭代器和如下遍历，其他遍历可能解析报错）
        ArrayList<List<String>> rows = new ArrayList<>();
        for (Row row : sheet) {
            StringBuilder line = new StringBuilder();
            ArrayList<String> cells = new ArrayList<>();
            for (Cell cell : row) {
                line.append(cell.toString().trim());
                cells.add(cell.toString().trim());
            }
            if (line.toString().trim().length() > 0){
                rows.add(cells);
            }
        }
        return rows;
    }

    // 释放资源
    public static void excelDestroy() throws IOException {
        workbook.close();
        excel.close();
    }

    // Excel解析后的字符串写入文本文件
    public static void excelWriter(String text, String file, Boolean append) throws IOException {
        // 合法性判断
        File f = new File(file);
        String fileName = f.getName();
        if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
            System.out.println("Failed to write the file. InvalidException: The file type is invalid.");
            return;
        }
        // 不存在创建
        if (!f.exists()) {
            System.out.println(f.createNewFile() ? "File created successfully.": "File creation failure.");
        }
        // 缓冲流
        BufferedWriter bw = new BufferedWriter(new FileWriter(file, append));
        // 写入字符串
        bw.write(text);
        // 刷新缓冲区
        bw.flush();
        bw.close();
        System.out.println("File write success.");
    }


}

