package com.example.reportdownload.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
public class htmlController {

    @RequestMapping("/download")
    public String toLogin(){
        return "download";
    }
}
