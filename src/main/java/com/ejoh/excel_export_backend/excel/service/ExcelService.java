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
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@Service
public class ExcelService {

    public static List<String> headerList = null;
    public static List<String> etcList = null;
    public static String pickDate = "";

    public static Map<String, List<Map<String, String>>> resultMap = null;

    /**
     * 
     * https://europani.github.io/spring/2022/01/05/035-excel.html 참고
     * https://code-killer.tistory.com/133
     * 
     * @param response
     * @param params
     * @return
     * @throws IOException
     */
    public ResponseEntity<Object> downExcel(HttpServletResponse response, Map<String, String> params) throws Exception {
        log.info("====================== Service start ======================");

        // for(String p : params.keySet()) {
        //     log.info(p + " : " + params.get(p));
        // }

        // Excel 파일 생성
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        int result = makeExcel(out, params); // 파일 저장 없이 Stream으로 생성

        // 헤더 설정
        HttpHeaders header = new HttpHeaders();
        header.add(HttpHeaders.CONTENT_DISPOSITION, "attachment");
        header.add("Cache-Control", "no-cache, no-store, must-revalidate");
        header.add("Pragma", "no-cache");
        header.add("Expires", "0");

        // 바디 설정
        ByteArrayResource resource = new ByteArrayResource(out.toByteArray());
        
        log.info("====================== Service end ======================");

        if(result > 1) {
            return ResponseEntity.ok()
                                .headers(header)        // 헤더 삽입
                                .contentLength(out.size())
                                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                                .body(resource);        // 바디 삽입
        } else if(result == 0) {
            return ResponseEntity.ok()
                                .contentType(MediaType.TEXT_PLAIN)
                                .body("엑셀 파일 생성 실패. <br>경로를 확인해주세요.");
        } else if(result == 1) {
            return ResponseEntity.ok()
                                .contentType(MediaType.TEXT_PLAIN)
                                .body("엑셀 파일 생성 실패. <br>선택하신 확장자 또는 수정일을 확인해주세요.");
        } else {
            return null;
        }
    }

    public int makeExcel(ByteArrayOutputStream out, Map<String, String> params) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        StringTokenizer st;

        // 초기화
        headerList = new ArrayList<>();
        etcList = new ArrayList<>();
        pickDate = "";
        resultMap = new LinkedHashMap<>();

        // 시트 생성
        Sheet sheet = wb.createSheet("이행목록");

        // row(행) 와 cell(열) 선언 및 초기화
        int rowCnt = 0;
        int cellCnt = 0;

        // Header 설정
        Row headerRow = sheet.createRow(rowCnt++);
        st = new StringTokenizer(params.get("headerList"), "/");
        String header = "";
        while (st.hasMoreTokens()) {
            header = st.nextToken();
            headerRow.createCell(cellCnt++).setCellValue(header);
            headerList.add(header);
        }
        
        // Body 설정
        // dirPath  etcList  pickDate  headerList
        
        st = new StringTokenizer(params.get("etcList"), "/");
        while (st.hasMoreTokens()) {
            etcList.add(st.nextToken());
        }

        if(params.get("pickDate") != null) {
            pickDate = params.get("pickDate");
        }


        // 1. 전달받은 절대경로의 폴더 내 파일 리스트 읽기
        int result = dirFileList(params.get("dirPath"), rowCnt++, sheet);

        if(result > 1) {
            // 엑셀 Body 데이터 출력
            for(String No : resultMap.keySet()) {
                List<Map<String, String>> dataList = resultMap.get(No);
                for(Map<String, String> mapData : dataList) {
                    Row bodyRow = sheet.createRow(Integer.parseInt(No));
                    int cell = 0;
                    for(String key : mapData.keySet()) {
                        bodyRow.createCell(cell++).setCellValue(mapData.get(key));
                    }
                }
            }

            try {
                wb.write(out);
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                wb.close();
            }
        }

        return result;
    }

    /**
     * 특정 폴더의 파일 리스트 출력
     * @param path
     * @throws Exception 
     */
    public static int dirFileList(String path, int rowCnt, Sheet sheet) throws Exception {
        File dirPath = new File(path);
        File files[] = dirPath.listFiles();
        
        Path paths = Paths.get(path);
        if(Files.exists(paths) && files != null) {
            for(File f : files) {
                if(f.isDirectory()) {
                    rowCnt = dirFileList(f.getPath(), rowCnt, sheet);
                } else if(f.isFile()) {
                    // 2. 전달받은 확장자에 해당하는 파일 리스트 필터링
                    if(etcList.contains(FilenameUtils.getExtension(f.getName()))) {
                        List<Map<String, String>> resultList = new ArrayList<>();
                        Map<String, String> dataMap = new LinkedHashMap<>();

                        dataMap.put(headerList.get(0), String.valueOf(rowCnt));
                        dataMap.put(headerList.get(1), String.valueOf(f.getPath()));
                        dataMap.put(headerList.get(2), String.valueOf(f.getName()));

                        if(headerList.contains("확장자")) {
                            dataMap.put("확장자", FilenameUtils.getExtension(f.getName()));
                        }

                        // 3. 전달받은 날짜 이후에 수정된 파일 리스트 필터링
                        if(pickDate != null && pickDate.length() > 1 && headerList.contains("수정일")) {
                            BasicFileAttributes attr = Files.readAttributes(Paths.get(f.getPath()), BasicFileAttributes.class);
                            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

                            LocalDate lastModiDate = attr.lastModifiedTime().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                            LocalDate selDate = LocalDate.parse(pickDate, formatter);

                            if (selDate.isBefore(lastModiDate)) {
                                dataMap.put("수정일", lastModiDate.format(formatter));
                            }
                        }
                        resultList.add(dataMap);
                        if(headerList.size() == dataMap.size()) {
                            resultMap.put(String.valueOf(rowCnt), resultList);
                        }
                        rowCnt++;
                    }
                }
            }
        } else {
            return 0;
        }
        return rowCnt;
    }
}
