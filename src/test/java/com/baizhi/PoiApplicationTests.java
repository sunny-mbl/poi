package com.baizhi;

import com.baizhi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作簿创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");
        //通过工作表创建行
        HSSFRow row = sheet.createRow(0);
        //通过行创建单元格
        HSSFCell cell = row.createCell(0);
        //给单元格赋值
        cell.setCellValue("第一个单元格");

        //把这个文件导出
        try {
            workbook.write(new FileOutputStream(new File("F:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void test1() {
        //创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //通过工作簿创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");
        //设置单元格宽度
        sheet.setColumnWidth(2, 15 * 256);

        //设置日期格式
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat("yyyy年mm月dd号");
        //把日期格式交给样式对象
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format);

        //创建单元格样式对象
        HSSFCellStyle fontStyle = workbook.createCellStyle();
        fontStyle.setAlignment(HorizontalAlignment.CENTER);
        //创建字体样式对象
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(Font.COLOR_RED);
        font.setItalic(true);
        font.setFontName("华文行楷");
        fontStyle.setFont(font);

        //创建标题行
        HSSFRow tiltRow = sheet.createRow(0);
        String[] str = {"id", "姓名", "生日"};
        for (int i = 0; i < str.length; i++) {
            HSSFCell cell = tiltRow.createCell(i);
            cell.setCellStyle(fontStyle);
            cell.setCellValue(str[i]);
        }
        //填充内容
        User user = new User("1", "小白", new Date());
        User user2 = new User("2", "小黑", new Date());
        User user3 = new User("3", "小红", new Date());
        ArrayList<User> users = new ArrayList<>();
        users.add(user);
        users.add(user2);
        users.add(user3);
        for (int i = 0; i < users.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(users.get(i).getId());
            row.createCell(1).setCellValue(users.get(i).getName());
            HSSFCell cell = row.createCell(2);
            //将样式对象设置到当前单元格中
            cell.setCellStyle(cellStyle);
            cell.setCellValue(users.get(i).getBir());
        }
        try {
            workbook.write(new FileOutputStream(new File("F:/a.xls")));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void poiIn() {
        HSSFWorkbook workbook = null;
        try {
            workbook = new HSSFWorkbook(new FileInputStream(new File("D:a.xls")));
            HSSFSheet sheet = workbook.getSheet("测试");
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i <= lastRowNum; i++) {
                User user = new User();
                HSSFRow row = sheet.getRow(i);
                HSSFCell id = row.getCell(0);
                HSSFCell name = row.getCell(1);
                HSSFCell bir = row.getCell(2);
                user.setId(id.toString());
                user.setName(name.getStringCellValue());
                user.setBir(bir.getDateCellValue());
                System.out.println(user);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
