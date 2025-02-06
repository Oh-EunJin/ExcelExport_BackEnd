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

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ExcelService {

    /**
     * 엑셀 생성 함수
     * @param response
     * @return
     * @throws IOException 
     */
    public void createExcel(HttpServletResponse response, Map<String, String> params) throws IOException {
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
        headerRow.createCell(1).setCellValue("경로");
        headerRow.createCell(2).setCellValue("파일명");

        // Body 설정
        for(int i = 1; i <= 3; i++) {
            Row bodyRow = sheet.createRow(rowCnt++);
            bodyRow.createCell(0).setCellValue(i);
            bodyRow.createCell(1).setCellValue("경로" + i);
            bodyRow.createCell(2).setCellValue("file" + i);
        }

        // excel 파일 다운로드
        // content 타입 및 파일명 지정
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment");

        try {
            wb.write(response.getOutputStream());
        } catch (Exception e) {
            e.getMessage();
        } finally {
            wb.close();
        }
        
        log.info("====================== Service end ======================");
    }

    /**
     * 
     * https://europani.github.io/spring/2022/01/05/035-excel.html 참고
     * https://code-killer.tistory.com/133
     * @param response
     * @param params
     * @return
     * @throws IOException
     */
    public ResponseEntity<Resource> createExcel2(HttpServletResponse response, Map<String, String> params) throws IOException {
        log.info("====================== Service start ======================");

        // String fileName = "일단임시이름" + ".xlsx";
        // File file = new File(fileName);

        // Excel 파일 생성
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        makeExcel(out); // 파일 저장 없이 Stream으로 생성
        // makeExcel(file);

        // 헤더 설정
        HttpHeaders header = new HttpHeaders();
        header.add(HttpHeaders.CONTENT_DISPOSITION, "attachment");
        header.add("Cache-Control", "no-cache, no-store, must-revalidate");
        header.add("Pragma", "no-cache");
        header.add("Expires", "0");

        // 바디 설정
        // InputStreamResource resource = new InputStreamResource(new FileInputStream(file));
        ByteArrayResource resource = new ByteArrayResource(out.toByteArray());
        
        log.info("====================== Service end ======================");

        return ResponseEntity.ok()
                            .headers(header)        // 헤더 삽입
                            .contentLength(out.size())
                            .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                            .body(resource);        // 바디 삽입
    }

    public void makeExcel(ByteArrayOutputStream out /* File file */) throws IOException {

        XSSFWorkbook wb = new XSSFWorkbook();

        // 시트 생성
        Sheet sheet = wb.createSheet("이행목록");

        // row 와 cell 선언 및 초기화
        int rowCnt = 0;
        int cellCnt = 0;

        // Header 설정
        Row headerRow = sheet.createRow(rowCnt++);
        headerRow.createCell(0).setCellValue("No");
        headerRow.createCell(1).setCellValue("경로");
        headerRow.createCell(2).setCellValue("파일명");

        // Body 설정
        for(int i = 1; i <= 3; i++) {
            Row bodyRow = sheet.createRow(rowCnt++);
            bodyRow.createCell(0).setCellValue(i);
            bodyRow.createCell(1).setCellValue("경로" + i);
            bodyRow.createCell(2).setCellValue("file" + i);
        }

        try {
            wb.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            wb.close();
        }
    }
}
