package com.ejoh.excel_export_backend.excel.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ejoh.excel_export_backend.excel.service.ExcelService;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@RestController
@RequestMapping("/api")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/test")
    public String Test() {
        
        log.info("dummyyyyyyyy");
        return "Vue와 연동 테스트";
    }
}
