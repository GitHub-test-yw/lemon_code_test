package com.four13;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

/**
 * @author NaTie
 * @date 2021/4/14 - 0:48
 * 学习java自动化测试
 */
public class ExcelDemo {
    public static void main(String[] args) throws Exception {
        /*
        *1、创建输入流，加载excel文件
        *2、调用POI API创建EXCEL对象
        *3、获取sheet
        *4、获取row（列）
        *5、获取cell（行）
        *6、输出内容
        * */
        // 1、创建输入流，加载excel文件
        FileInputStream fis = new FileInputStream("src/test/resources/java28.xls");
        // 2、调用POI API创建EXCEL对象
        Workbook sheets = WorkbookFactory.create(fis); //WorKbook都是接口，是多态
        //关流
        fis.close();//这里关流是因为create中的fis已经将数据读取到sheets中了
        // 3、获取sheet，Sheet是接口，是多态
        Sheet sheet = sheets.getSheetAt(0);//根据索引获取，获取也可以直接写getSheet方法
        // 4、获取row（行），Row是接口
        Row row =sheet.getRow(3);
        // 5、获取cell（单元格），Cell是接口
        Cell cell = row.getCell(1);
        // 6、输出内容
        System.out.println(cell.getStringCellValue());

    }
}
