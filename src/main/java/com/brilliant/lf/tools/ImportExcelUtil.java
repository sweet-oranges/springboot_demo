package com.brilliant.lf.tools;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class ImportExcelUtil {

    private final static String excel2003L =".xls";    //2003- 版本的excel
    private final static String excel2007U =".xlsx";   //2007+ 版本的excel

    public final static List<Object> finalList= new ArrayList<>();
    public final static Map<String,List<Object>> finalMap=new HashMap<>(9);

    /**
     * 描述：获取IO流中的数据，组装成List<List<Object>>对象
     * @param in,fileName
     * @return
     * @throws IOException
     */
    public List<List<Object>> getBankListByExcel(InputStream in, String fileName) throws Exception {
        List<List<Object>> list = null;

        //创建Excel工作薄
        Workbook work = this.getWorkbook(in,fileName);
        if(null == work){
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        list = new ArrayList<List<Object>>();
        //遍历Excel中所有的sheet
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if(sheet==null){continue;}

            //遍历当前sheet中的所有行
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if(row==null||row.getFirstCellNum()==j){continue;}

                //遍历所有的列
                List<Object> li = null;
                try {
                    li = new ArrayList<Object>();
                    for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                        int x = row.getLastCellNum();
                        cell = row.getCell(y);
                        if (cell == null){
                            li.add("");
                        }else{
                            li.add(this.getCellValue(cell));
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    throw new Exception("导入错误，存在数据为空");
                }
                list.add(li);
            }
        }
        return list;
    }

    /**
     * 描述：获取IO流中的数据，组装成List<List<List<Object>>>对象
     * @param in,fileName
     * @return
     * @throws IOException
     */
    public List<List<List<Object>>> getMapListByExcel(InputStream in, String fileName) throws Exception {
        List<List<List<Object>>> list = null;

        //创建Excel工作薄
        Workbook work = this.getWorkbook(in,fileName);
        if(null == work){
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;

        list = new ArrayList<List<List<Object>>>();
        //遍历Excel中所有的sheet
        List<List<Object>> lis = null;
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            lis = new ArrayList<List<Object>>();
            sheet = work.getSheetAt(i);
            if(sheet==null){continue;}

            //遍历当前sheet中的所有行
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if(row==null||row.getFirstCellNum()==j){continue;}

                //遍历所有的列
                List<Object> li = null;
                try {
                    li = new ArrayList<Object>();
                    for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                        int x = row.getLastCellNum();
                        cell = row.getCell(y);
                        if (cell == null){
                            li.add("");
                        }else{
                            li.add(this.getCellValue(cell));
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                    throw new Exception("导入错误，存在数据为空");
                }
                switch (i){
                    case 0:
                        lis.add(li);
                        break;
                    case 1:
                        lis.add(li);
                        break;
                    case 2:
                        lis.add(li);
                        break;
                    case 3:
                        lis.add(li);
                        break;
                }
            }
            list.add(lis);
        }
        return list;
    }

    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * @param inStr,fileName
     * @return
     * @throws Exception
     */
    public Workbook getWorkbook(InputStream inStr, String fileName) throws Exception {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if(excel2003L.equals(fileType)){
            wb = new HSSFWorkbook(inStr);  //2003-
        }else if(excel2007U.equals(fileType)){
            wb = new XSSFWorkbook(inStr);  //2007+
        }else{
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     * 描述：对表格中数值进行格式化
     * @param cell
     * @return
     */
    public Object getCellValue(Cell cell){
        Object value = null;
        DecimalFormat df = new DecimalFormat("0");  //格式化number String字符
        SimpleDateFormat sdf = new SimpleDateFormat("yyy-MM-dd hh:mm:ss");  //日期格式化
        DecimalFormat df2 = new DecimalFormat("0.00");  //格式化数字

        switch (cell.getCellType()) {
            case STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if("General".equals(cell.getCellStyle().getDataFormatString())){
                    value = df.format(cell.getNumericCellValue());
                }else if(cell.toString().contains("-") && checkDate(cell.toString())){
                    value = sdf.format(cell.getDateCellValue());
                }else{
                    value = df2.format(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case BLANK:
                value = "";
                break;
            default:
                break;
        }
        return value;
    }

    /**
     *  判断是否是“02-十一月-2006”格式的日期类型
     */
     private static boolean checkDate(String str){
         String[] dataArr =str.split("-");
         try {
             if(dataArr.length == 3){
                 int x = Integer.parseInt(dataArr[0]);
                 String y =  dataArr[1];
                 int z = Integer.parseInt(dataArr[2]);
                 if(x>0 && x<32 && z>0 && z< 10000 && y.endsWith("月")){
                     return true;
                 }
             }
         } catch (Exception e) {
             return false;
         }
         return false;
     }

    public static void main(String[] args){
        try {
            InputStream is = new FileInputStream("d:\\test.xls");
            List<List<List<Object>>> list = new ImportExcelUtil().getMapListByExcel(is,"test.xls");

            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}