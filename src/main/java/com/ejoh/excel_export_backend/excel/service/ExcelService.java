/**
 * packageName    : com.ejoh.excel_export_backend.excel.service
 * fileName       : ExcelService
 * author         : 오은진
 * date           : 2025.01.22
 * description    : 엑셀 생성을 위한 Service
 * ===========================================================
 * DATE              AUTHOR            NOTE
 * -----------------------------------------------------------
 * 2025.01.22        오은진            엑셀 파일 생성 코드 임시 작성
 */

package com.ejoh.excel_export_backend.excel.service;

import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ExcelService {
    
    public String dummy() {
        return "NoteBoooook!!!!";
    }

    /**
     * 엑셀 생성 함수
     * @param response
     * @return
     */
    public void createExcel(HttpServletResponse response, Map<String, String> params) {
        log.info("====================== Service start ======================");
        // for(String p : params.keySet()) {
        //     log.info(params.get(p));
        // } 

        XSSFWorkbook wb = new XSSFWorkbook();

        // 시트 생성
        Sheet sheet = wb.createSheet("이행목록");

        // row 와 cell 선언 및 초기화
        int rowCnt = 0;
        int cellCnt = 0;

        // Header 설정
        Row headerRow = sheet.createRow(rowCnt++);
        headerRow.createCell(0).setCellValue("No");
        headerRow.createCell(1).setCellValue("파일명");
        headerRow.createCell(2).setCellValue("경로");

        // Body 설정
        for(int i = 1; i <= 3; i++) {
            Row bodyRow = sheet.createRow(rowCnt++);
            bodyRow.createCell(0).setCellValue(i);
            bodyRow.createCell(1).setCellValue("file" + i);
            bodyRow.createCell(2).setCellValue("경로" + i);
        }

        // excel 파일 다운로드
        // content 타입 및 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment");

        try {
            wb.write(response.getOutputStream());
            wb.close();
        } catch (Exception e) {
            e.getMessage();
        }
        
        log.info("====================== Service end ======================");
    }
}
