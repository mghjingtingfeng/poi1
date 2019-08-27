package com.baizhi.poi;

import com.baizhi.poi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {


    @Test
    public void contextLoads() throws IOException {
        HSSFWorkbook sheets = new HSSFWorkbook();
        HSSFSheet sheet = sheets.createSheet("测试");
        HSSFRow row = sheet.createRow(0);
        String[] str = {"id", "name", "bir"};
        for (int i = 0; i < str.length; i++) {
            row.createCell(i).setCellValue(str[i]);
        }

        sheets.write(new FileOutputStream(new File("D:/a.xls")));
    }

    @Test
    public void test1() throws IOException {
        HSSFWorkbook sheets = new HSSFWorkbook();
        HSSFSheet sheet = sheets.createSheet("测试");
        HSSFRow row = sheet.createRow(0);

        HSSFDataFormat dataFormat = sheets.createDataFormat();
        short format = dataFormat.getFormat("yyyy-MM-dd");

        HSSFCellStyle cellStyle = sheets.createCellStyle();
        cellStyle.setDataFormat(format);

        HSSFFont font = sheets.createFont();
        font.setBold(true);
        font.setItalic(true);
        font.setColor(Font.COLOR_RED);

        HSSFCellStyle style = sheets.createCellStyle();
        style.setFont(font);

        String[] str = {"id", "name", "bir"};
        for (int i = 0; i < str.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(str[i]);
        }


        User user = new User("1", "小黑", new Date());
        User user2 = new User("2", "小兰", new Date());
        User user3 = new User("3", "小李", new Date());

        ArrayList<User> users = new ArrayList<>();
        users.add(user);
        users.add(user2);
        users.add(user3);
        for (int i = 0; i < users.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            row1.createCell(0).setCellValue(users.get(i).getId());
            row1.createCell(1).setCellValue(users.get(i).getName());
            HSSFCell cell = row1.createCell(2);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(users.get(i).getBir());
        }
        sheet.setColumnWidth(2, 15);


        sheets.write(new FileOutputStream(new File("D:/a.xls")));
    }

    @Test
    public void test2() {
        System.out.println("hello");
        System.out.println("world");
        System.out.println("xiaoheihei");
        System.out.println("牛批");
        System.out.println("hi");
        System.out.println("傻叉");
        System.out.println("傻屌");
        System.out.println("金雕");
    }

}
