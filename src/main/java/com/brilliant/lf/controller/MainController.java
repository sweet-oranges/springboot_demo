package com.brilliant.lf.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * 主要控制类
 *
 * @Author zxl on 2019/11/4
 */
@Controller
public class MainController {

    @RequestMapping("/")
    public String hello(){
        return "dad";
    }
}