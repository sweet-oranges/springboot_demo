package com.brilliant.lf.controller;


import org.apache.tomcat.util.ExceptionUtils;
import org.apache.tomcat.util.http.fileupload.IOUtils;
import org.apache.tomcat.util.http.fileupload.util.Streams;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Calendar;
import java.util.UUID;


/**
 * @description: 上传控制类
 * @Author: zxl on 2020/2/29
 * @create: 2020-02-29 09:26
 */
@RestController
@CrossOrigin
@RequestMapping("/api/upload")
public class UploadController {

    private static final Logger logger = LoggerFactory.getLogger(UploadController.class);
    //在文件操作中，不用/或者\最好，推荐使用File.separator
    private final static String  fileDir="files";
    private final static String rootPath = "D:"+File.separator+fileDir+File.separator;

    @RequestMapping("/uploadPic")
    public String uploadPic(@RequestParam(value = "file",required = false)MultipartFile file, HttpServletRequest request){
        //写死的本地硬盘
        String path = "D:/img";
        //获取文件名称
        String fileName = file.getOriginalFilename();
        //获取
        Calendar currTime = Calendar.getInstance();
        String time = String.valueOf(currTime.get(Calendar.YEAR))+String.valueOf((currTime.get(Calendar.MONTH)+1));
        //获取文件名后缀
        String suffix = fileName.substring(file.getOriginalFilename().lastIndexOf("."));
        suffix = suffix.toLowerCase();
        if(suffix.equals(".jpg") || suffix.equals("jpeg") || suffix.equals("png")){
            fileName = UUID.randomUUID().toString()+suffix;
            File targetFile = new File(path,fileName);
            if(!targetFile.getParentFile().exists()){
                targetFile.getParentFile().mkdirs();
            }
            long size = 0;
            //保存
            try{
                file.transferTo(targetFile);
                size = file.getSize();
            }catch(Exception e){
                e.printStackTrace();
                return "上传失败";
            }
            //项目地址
            String fileUrl = "http://localhost:8888";
            //文件获取地址
            fileUrl = fileUrl + request.getContextPath()+"/img/"+fileName;
            return fileUrl;

        }else{
            return "上传图片格式有误";
        }
    }

    @RequestMapping("/upload2")
    public Object uploadFile(@RequestParam("file")MultipartFile[] multipartFiles, final HttpServletResponse response,final HttpServletRequest request){
        File fileDir = new File(rootPath);
        if(!fileDir.exists() && !fileDir.isDirectory()){
            fileDir.mkdirs();
        }
        try{
            if(multipartFiles != null && multipartFiles.length > 0){
                for(int i = 0; i < multipartFiles.length; i++){
                    try{
                        //以原来的名称命名，覆盖掉旧的
                        String storagePath = rootPath+multipartFiles[i].getOriginalFilename();
                        logger.info("上传的文件:"+multipartFiles[i].getName()+","+multipartFiles[i].getContentType()+","+multipartFiles[i].getOriginalFilename()
                        +",保存的路径为："+storagePath);
                        Streams.copy(multipartFiles[i].getInputStream(),new FileOutputStream(storagePath),true);
                        //或者下面的
                        Path path = Paths.get(storagePath);
                        Files.write(path,multipartFiles[i].getBytes());
                    }catch (IOException e){
                        logger.error("错误");
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }
        return "上传成功";
    }

    @RequestMapping("/download")
    public Object downloadFile(final HttpServletResponse response, final HttpServletRequest request){
        String fileName = "2测试文档.doc";
        OutputStream os = null;
        InputStream is= null;
        try{
            // 取得输出流
            os = response.getOutputStream();
            // 清空输出流
            response.reset();
            response.setContentType("application/x-download;charset=GBK");
            response.setHeader("Content-Disposition", "attachment;filename="+ new String(fileName.getBytes("utf-8"), "iso-8859-1"));
            //读取流
            File f = new File(rootPath+fileName);
            is = new FileInputStream(f);
            if (is == null) {
                logger.error("下载附件失败，请检查文件“" + fileName + "”是否存在");
                return "下载附件失败，请检查文件“" + fileName + "”是否存在";
            }
            //复制
            IOUtils.copy(is, response.getOutputStream());
            response.getOutputStream().flush();
        }catch (IOException e){
            return "下载附件失败,error:"+e.getMessage();

        }
        //文件的关闭放在finally中
        finally
        {
            try {
                if (is != null) {
                    is.close();
                }
            } catch (IOException e) {

            }
            try {
                if (os != null) {
                    os.close();
                }
            } catch (IOException e) {

            }
        }
        return null;
    }
}