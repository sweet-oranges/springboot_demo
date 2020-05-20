package com.brilliant.lf.tools;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * @description:
 * @Author: zxl on 2020/5/20
 * @create: 2020-05-20 08:40
 */
public class ImportExcel {
    private POIFSFileSystem fs;
    private HSSFWorkbook wb;
    private HSSFSheet sheet;
    private HSSFRow row;

    /**
     27      * 读取Excel表格表头的内容
     28      * @param is
     29      * @return String 表头内容的数组
     30      */
      public String[] readExcelTitle(InputStream is) {
                 try {
                         fs = new POIFSFileSystem(is);
                         wb = new HSSFWorkbook(fs);
                     } catch (IOException e) {
                         e.printStackTrace();
                     }
                 sheet = wb.getSheetAt(0);
                 //得到首行的row
                 row = sheet.getRow(0);
                 // 标题总列数
                 int colNum = row.getPhysicalNumberOfCells();
                 String[] title = new String[colNum];
                 for (int i = 0; i < colNum; i++) {
                         title[i] = getCellFormatValue(row.getCell((short) i));
                     }
                 return title;
             }

              /**
  51      * 读取Excel数据内容
  52      * @param is
  53      * @return Map 包含单元格数据内容的Map对象
  54      */
              public Map<Integer, String> readExcelContent(InputStream is) {
                 Map<Integer, String> content = new HashMap<Integer, String>();
                 String str = "";
                 try {
                         fs = new POIFSFileSystem(is);
                         wb = new HSSFWorkbook(fs);
                     } catch (IOException e) {
                         e.printStackTrace();
                     }
                 sheet = wb.getSheetAt(0);
                 // 得到总行数
                 int rowNum = sheet.getLastRowNum();
                 //由于第0行和第一行已经合并了  在这里索引从2开始
                 row = sheet.getRow(2);
                 int colNum = row.getPhysicalNumberOfCells();
                 // 正文内容应该从第二行开始,第一行为表头的标题
                 for (int i = 2; i <= rowNum; i++) {
                         row = sheet.getRow(i);
                         int j = 0;
                         while (j < colNum) {
                                 str += getCellFormatValue(row.getCell((short) j)).trim() + "-";
                                 j++;
                             }
                         content.put(i, str);
                         str = "";
                     }
                 return content;
             }

              /**
  85      * 获取单元格数据内容为字符串类型的数据
  86      *
  87      * @param cell Excel单元格
  88      * @return String 单元格数据内容
  89      */
              private String getStringCellValue(HSSFCell cell) {
                 String strCell = "";
                 switch (cell.getCellType()) {
                     case STRING:
                                 strCell = cell.getStringCellValue();
                                 break;
                         case NUMERIC:
                                 strCell = String.valueOf(cell.getNumericCellValue());
                                 break;
                         case BOOLEAN:
                                 strCell = String.valueOf(cell.getBooleanCellValue());
                                 break;
                        case BLANK:
                                 strCell = "";
                                 break;
                         default:
                                 strCell = "";
                                 break;
                     }
                 if (strCell.equals("") || strCell == null) {
                         return "";
                     }
                 if (cell == null) {
                         return "";
                     }
                 return strCell;
             }

             /**
 119      * 获取单元格数据内容为日期类型的数据
 120      *
 121      * @param cell
 122      *            Excel单元格
 123      * @return String 单元格数据内容
 124      */
             private String getDateCellValue(HSSFCell cell) {
                 String result = "";
                 try {
                         CellType cellType = cell.getCellType();
                         if (cellType == CellType.NUMERIC) {
                                 Date date = cell.getDateCellValue();
                                 result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1)
                                         + "-" + date.getDate();
                             } else if (cellType == CellType.STRING) {
                                 String date = getStringCellValue(cell);
                                 result = date.replaceAll("[年月]", "-").replace("日", "").trim();
                             } else if (cellType == CellType.BLANK) {
                                 result = "";
                             }
                     } catch (Exception e) {
                         System.out.println("日期格式不正确!");
                         e.printStackTrace();
                     }
                 return result;
             }

             /**
 147      * 根据HSSFCell类型设置数据
 148      * @param cell
 149      * @return
 150      */
             private String getCellFormatValue(HSSFCell cell) {
                 String cellvalue = "";
                 if (cell != null) {
                         // 判断当前Cell的Type
                         switch (cell.getCellType()) {
                                 // 如果当前Cell的Type为NUMERIC
                                 case NUMERIC:
                                     case FORMULA: {
                                         // 判断当前的cell是否为Date
                                         if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                                 Date date = cell.getDateCellValue();
                                                 SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                                                 cellvalue = sdf.format(date);
                                             }
                                         // 如果是纯数字
                                         else {
                                                 // 取得当前Cell的数值
                                                 cellvalue = String.valueOf(cell.getNumericCellValue());
                                             }
                                         break;
                                     }
                                 // 如果当前Cell的Type为STRIN
                                 case STRING:
                                         // 取得当前的Cell字符串
                                         cellvalue = cell.getRichStringCellValue().getString();
                                         break;
                                 // 默认的Cell值
                                 default:
                                         cellvalue = " ";
                                 }
                     } else {
                        cellvalue = "";
                    }
                 return cellvalue;

             }

    public static void main(String[] args) {
                 try {
                         // 对读取Excel表格标题测试
                         InputStream is = new FileInputStream("d:\\test.xls");
                         ImportExcel excelReader = new ImportExcel();
                         String[] title = excelReader.readExcelTitle(is);
                         System.out.println("获得Excel表格的标题:");
                         for (String s : title) {
                                 System.out.print(s + " ");
                             }
                         System.out.println();

                         // 对读取Excel表格内容测试
                         InputStream is2 = new FileInputStream("d:\\test.xls");
                         Map<Integer, String> map = excelReader.readExcelContent(is2);
                         System.out.println("获得Excel表格的内容:");
                         //这里由于xls合并了单元格需要对索引特殊处理
                         for (int i = 2; i <= map.size()+1; i++) {
                                 System.out.println(map.get(i));
                             }

                     } catch (FileNotFoundException e) {
                         System.out.println("未找到指定路径的文件!");
                         e.printStackTrace();
                     }
             }
}