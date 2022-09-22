package com.example.excelfileread;


import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Slf4j
public class ExcelRead {

    public static List<Map<String, String>> read(ReadOption readOption) {

        Workbook wb = FileType.getWorkbook(readOption.getFilePath());

        Sheet sheet = wb.getSheetAt(0);

        int numOfRows = sheet.getPhysicalNumberOfRows();
        int numOfCells = 0;

        Row row = null;
        Cell cell = null;

        String cellName = "";

        Map<String,String> map = null;

        List<Map<String,String>> result = new ArrayList<Map<String,String>>();

        for(int rowIndex = readOption.getStartRow() -1; rowIndex <numOfRows; rowIndex++){

            row = sheet.getRow(rowIndex);

            if(row != null){

                numOfCells = row.getPhysicalNumberOfCells();

                map = new HashMap<String, String>();

                for(int cellIndex = 0; cellIndex < numOfCells; cellIndex++){

                    cell = row.getCell(cellIndex);

                    cellName = CellRef.getName(cell, cellIndex);

                    if(!readOption.getOutputColumns().contains(cellName)){
                        continue;
                    }

                    map.put(cellName, CellRef.getValue(cell));
                }

                result.add(map);
            }

        }

        return result;

    }

    public static void main(String[] args) {

        ReadOption ro = new ReadOption();
        ro.setFilePath("/Users/taehyeonkim/Desktop/incuvers/통합고지 가스전기데이터/가스 계약자 DB.xlsx");
        // C : 고객명 , N : 호수 , O : 지번주소, H : 시/도, I : 구/군, J : 동
        ro.setOutputColumns("C","H","I","J","K","N","O");
        ro.setStartRow(2);

        List<Map<String,String>> result = ExcelRead.read(ro);

        System.out.println(ro);

        for(Map<String,String> map : result){
            String clNm = map.get("C");
            String sido = map.get("H");
            String gugun = map.get("I");
            String dong = map.get("J");
            // 지번 주소
            String jibun = map.get("K");

            //호수
            String hosu = map.get("N");

            StringBuffer dongjibun = new StringBuffer();
            dongjibun.append(sido+" "+gugun+" "+dong);

            String ad = dongjibun.toString();

            try {

                // 주소 검색, URLEncoder는 URL을 인코딩 하기위해 사용하는 클래스
                String keyword = URLEncoder.encode(ad, "UTF-8");

                URL url = new URL("https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword="
                        + keyword + "&confmKey=devU01TX0FVVEgyMDIyMDkwMjE2MTYzOTExMjk0NDM=&resultType=json");

                BufferedReader bf;


                bf = new BufferedReader(new InputStreamReader(url.openStream(), "UTF-8"));

                String res = bf.readLine();

//                System.out.println(res);

                // String 값을 JSON 형태로 추출하기 위해 사용하는 라이브러리
                JSONParser jsonParser = new JSONParser();
                JSONObject jsonObject = (JSONObject)jsonParser.parse(res);
                JSONObject addResult = (JSONObject)jsonObject.get("results");

                // 리스트 추출
                JSONArray jusoArray = (JSONArray)addResult.get("juso");

                // 컬렉션 추출 주소정보 뽑을 준비 완료!
                JSONObject jusoColl = (JSONObject) jusoArray.get(0);

                // 행정동 코드
                String hCode = jusoColl.get("admCd").toString();              // 행정동코드 10자리

                // 숫자 빼고 모두 제거
                String match = "[^0-9]";
                jibun = jibun.replaceAll(match, "");
                hosu = hosu.replaceAll(match, "");
                int intJibun = Integer.parseInt(jibun);                       // int 로 형변환
                int intHosu = Integer.parseInt(hosu);

                String jibunNm = String.format("%06d",intJibun);              // 최종 지번 숫자만(6자리)

                String hosuNm = String.format("%08d",intHosu);                // 최종 호수 숫자만(8자리)



                log.info("고객명: " + clNm + " | " + "n차 통합코드: " + hCode + jibunNm + hosuNm);

            } catch (Exception e) {
                log.error("잘못된 접근입니다",e);
            }
        }
    }




}
