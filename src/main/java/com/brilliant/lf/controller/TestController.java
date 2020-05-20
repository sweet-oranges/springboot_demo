package com.brilliant.lf.controller;

import com.brilliant.lf.bean.Person;
import com.brilliant.lf.service.PersonService;
import com.brilliant.lf.tools.ExcelUtils;
import com.brilliant.lf.tools.ImportExcelUtil;
import com.brilliant.lf.tools.POIUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.plaf.multi.MultiPopupMenuUI;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 学习springboot框架
 *
 * @Author zxl on 2019/8/9
 */
@RestController
@CrossOrigin
@RequestMapping("/api")
public class TestController {

    @Autowired
    private PersonService personService;


    @RequestMapping("/getAll")
    public void getAll(){
        System.out.println(personService.getAll());
        try {
            POIUtil.export(personService.getAll());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @RequestMapping("import")
    public void importExcel(){
        try {
           List<Person> list = POIUtil.importExcel();
            for (Person p:
                 list) {
                System.out.println(p.getName());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @RequestMapping("/ukey")
    public void ukey(@RequestParam String KeyName,@RequestParam String KeyPwds){
        System.out.println("用户名:"+KeyName+"---"+"密码:"+KeyPwds);
    }

    @RequestMapping("/login")
    public void login(){

    }

    @RequestMapping("/get_info")
    public Map get_info(){
        Map<String,Object> map = new HashMap<>(10);
        map.put("code","200");
        map.put("avatar","http://localhost:8888/img/1d5b133e-0198-4e91-9f5f-f8840070f529.jpg");
        return map;
    }

    @RequestMapping("/exportData")
    public void exportUserFrom (HttpServletRequest request, HttpServletResponse response){
        try {
            List<Person> list = personService.getAll();
            String[] headerName = {"序号", "姓名"};
            String[] headerKey = {"id", "name"};
            HSSFWorkbook wb = ExcelUtils.createExcel(headerName, headerKey, "人员信息表", list);
            if (wb == null) {
                return;
            }
            response.setContentType("application/vnd.ms-excel");
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
            Date date = new Date();
            String str = sdf.format(date);
            String fileName = "数据" + str;
            response.setHeader("Content-disposition",
                    "attachment;filename=" + new String(fileName.getBytes("gb2312"), "ISO-8859-1") + ".xls");
            OutputStream outputStream = response.getOutputStream();
            outputStream.flush();
            wb.write(outputStream);
            outputStream.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    @RequestMapping("/test")
    public void test(){
        try {
            InputStream is = new FileInputStream("d:\\test.xls");
            List<List<List<Object>>> list = new ImportExcelUtil().getMapListByExcel(is,"test");

            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }




}