package com.brilliant.lf.tools;

import com.brilliant.lf.bean.Person;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
/**
 * @description:
 * @Author: zxl on 2020/5/19
 * @create: 2020-05-19 10:07
 */
public class POIUtil {
    public static void export(List<Person> personList) throws Exception{
        //指定数据存放的位置
        OutputStream outputStream = new FileOutputStream("D:\\test.xls");
        //1.创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //2.创建一个工作表sheet
        HSSFSheet sheet = workbook.createSheet("test");
        //构造参数依次表示起始行，截至行，起始列， 截至列
        CellRangeAddress region=new CellRangeAddress(0, 0, 0, 3);
        sheet.addMergedRegion(region);

        HSSFCellStyle style=workbook.createCellStyle();
        //水平居中
        style.setAlignment(HorizontalAlignment.CENTER);
        //垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell = row1.createCell(0);

        //设置值，这里合并单元格后相当于标题
        cell.setCellValue("人员信息表");

        //将样式添加生效
        cell.setCellStyle(style);

        for(int i = 0;i<personList.size();i++){
            //行
            HSSFRow row = sheet.createRow(i+1);
            //对列赋值
            row.createCell(0).setCellValue(personList.get(i).getId());
            row.createCell(1).setCellValue(personList.get(i).getName());
        }
        workbook.write(outputStream);
        outputStream.close();
    }
    public static List<Person> importExcel() throws Exception{
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("D:\\test.xls")));
        HSSFSheet sheet = null;
        List<Person> list = new ArrayList<Person>();
        for(int i = 0;i < workbook.getNumberOfSheets();i++){
            //获取每个sheet
            sheet = workbook.getSheetAt(i);

            //getPhysicalNumberOfRows获取有记录的行数
            for(int j = 0;j < sheet.getPhysicalNumberOfRows();j++){
                Row row = sheet.getRow(j);
                if(null!=row){
                    //getLastCellNum获取最后一列
                    Person user = new Person();
                    for(int k = 0;k < row.getLastCellNum();k++){
                        if(null!=row.getCell(k)){
                            if(k==0){
                                Cell cell = row.getCell(0);
                                //cell->double
                                Double d = cell.getNumericCellValue();
                                //double->int
                                int id = new Double(d).intValue();
                                user.setId(id);
                            }
                            if(k==1){
                                Cell cell = row.getCell(1);
                                //cell->string
                                user.setName(cell.getStringCellValue().toString());
                            }
                        }
                    }
                    list.add(user);
                }
            }
            System.out.println("读取sheet表："+ workbook.getSheetName(i) + "完成");

        }
        return list;
    }



}