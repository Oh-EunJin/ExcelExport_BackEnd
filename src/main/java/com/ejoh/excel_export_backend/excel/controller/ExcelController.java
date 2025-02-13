/**
 * packageName    : com.ejoh.excel_export_backend.excel.controller
 * fileName       : ExcelController
 * author         : 오은진
 * date           : 2025.01.22
 * description    : 엑셀 생성을 위한 Controller
 * ===========================================================
 * DATE              AUTHOR             NOTE
 * -----------------------------------------------------------
 * 2025.01.22        오은진            최초 생성
 */

package com.ejoh.excel_export_backend.excel.controller;

import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.ejoh.excel_export_backend.excel.service.ExcelService;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;


@Slf4j
@RestController
@RequestMapping("/api")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/excel")
    public ResponseEntity<Object> excelDownload(HttpServletResponse response, @RequestParam Map<String, String> params) throws Exception {

        log.info("====================== Controller start ======================");
        ResponseEntity<Object> responseEntity = excelService.excelDownload(response, params);
        log.info("====================== Controller end ======================");
        return responseEntity;
    }
    
}
