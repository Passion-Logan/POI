package com.cody.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * 应用模块名称<p>
 * 代码描述<p>
 * Copyright: Copyright (C) 2019 XXX, Inc. All rights reserved. <p>
 *
 * @author WQL
 * @since 2019年12月12日 0012 17:13
 */
public class demo {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("test");

        XSSFRow row;

        Map<String, Object[]> empinfo = new TreeMap<>();
        empinfo.put("1", new Object[]{
                "编号", "姓名", "称呼"});
        empinfo.put("2", new Object[]{
                "tp01", "Gopal", "Technical Manager"});
        empinfo.put("3", new Object[]{
                "tp02", "Manisha", "Proof Reader"});
        empinfo.put("4", new Object[]{
                "tp03", "Masthan", "Technical Writer"});
        empinfo.put("5", new Object[]{
                "tp04", "Satish", "Technical Writer"});
        empinfo.put("6", new Object[]{
                "tp05", "Krishna", "Technical Writer"});
        Set<String> keyid = empinfo.keySet();

        int rowid = 0;

        for (String key : keyid) {
            row = sheet.createRow(rowid++);
            Object[] objectArr = empinfo.get(key);
            int cellid = 0;
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }
        FileOutputStream out = new FileOutputStream(
                new File("Writesheet.xlsx"));
        workbook.write(out);
        out.close();
    }
}
