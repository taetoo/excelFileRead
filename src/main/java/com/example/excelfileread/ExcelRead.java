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
import java.util.*;

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
//        ro.setFilePath("/Users/taehyeonkim/Desktop/incuvers/통합고지 가스전기데이터/가스 계약자 DB.xlsx");
        ro.setFilePath("/Users/taehyeonkim/Desktop/무제 2.xlsx");
        // C : 고객명 , N : 호수 , O : 지번주소, H : 시/도, I : 구/군, J : 동
        ro.setOutputColumns("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X");
        ro.setStartRow(2);



        List<Map<String,String>> result = ExcelRead.read(ro);

        String hCode = "";

        for(Map<String,String> map : result){

            String bpNm = map.get("A");         // BP번호
            String caNm = map.get("B");         // CA번호
            String cliNm = map.get("C");        // 고객명
            String buildNm = map.get("D");      // 건물번호
            String houseNm = map.get("E");      // 세대번호
            String houseType = map.get("F");    // 세대유형
            String setNm = map.get("G");        // 설치번호
            String sido = map.get("H");         // 시/도
            String gugun = map.get("I");        // 구/군
            String dong = map.get("J");         // 동
            String jibun = map.get("K");        // 지번주소
            String buildName = map.get("L");    // 건물명
            String buildDong = map.get("M");    // 건물동
            String hosu = map.get("N");         // 호수
            String bJuso = map.get("O");        // 법정동주소
            String sinJuso = map.get("P");      // 신 주소
            String deliJuso = map.get("Q");     // 송달주소
            String deliNm = map.get("R");       // 송달번호
            String deliNmTxt = map.get("X");    // 송달번호TXT
            String readArea = map.get("T");     // 검침구역
            String readAreaTxt = map.get("U");  // 검침구역TXT
            String postNm = map.get("V");       // 우편번호
            String phoneNm = map.get("W");      // 전화번호
            String interCode = map.get("X");    // 통합코드

            // 건물 동의 값이 null 인 경우 & 숫자가 아닌 경우
            if(buildDong == null || buildDong == ""){
                buildDong = "0000";
            } else if(buildDong != "[0-9]") {
                buildDong = "0000";
            }


            // 숫자 빼고 모두 제거
            String match = "[^0-9]";
            jibun = jibun.replaceAll(match, "");
            buildDong = buildDong.replaceAll(match, "");
            hosu = hosu.replaceAll(match, "");
            // 호수가 숫자가 아닌경우
            if(hosu == "[^0-9]"){
                hosu = "0000";
            }


            int intJibun = Integer.parseInt(jibun);                       // int 로 형변환
            int intDongNm = Integer.parseInt(buildDong);
            int intHosu = Integer.parseInt(hosu);

            String jibunNm = String.format("%06d",intJibun);              // 최종 지번 숫자만(6자리)
            String dongNm = String.format("%04d",intDongNm);           // 최종 동 숫자만(4자리)
            String hosuNm = String.format("%04d",intHosu);           // 최종 동 숫자만(4자리)


            StringBuffer dongjibun = new StringBuffer();
            dongjibun.append(sido+" "+gugun+" "+dong);

            String addressName = dongjibun.toString();

            try {
                // 주소 검색, URLEncoder는 URL을 인코딩 하기위해 사용하는 클래스
                String keyword = URLEncoder.encode(addressName, "UTF-8");

                URL url = new URL("https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword="
                        + keyword + "&confmKey=devU01TX0FVVEgyMDIyMDkwMjE2MTYzOTExMjk0NDM=&resultType=json");

                BufferedReader bf;


                bf = new BufferedReader(new InputStreamReader(url.openStream(), "UTF-8"));

                String res = bf.readLine();

//              System.out.println(res);

                // String 값을 JSON 형태로 추출하기 위해 사용하는 라이브러리
                JSONParser jsonParser = new JSONParser();
                JSONObject jsonObject = (JSONObject)jsonParser.parse(res);
                JSONObject addResult = (JSONObject)jsonObject.get("results");

                // 리스트 추출
                JSONArray jusoArray = (JSONArray)addResult.get("juso");

                // 컬렉션 추출 주소정보 뽑을 준비 완료!
                JSONObject jusoColl = (JSONObject) jusoArray.get(0);

                // 행정동 코드
                hCode = jusoColl.get("admCd").toString();              // 행정동코드 10자리


            }

            catch (Exception e) {
                log.error("잘못된 접근입니다",e);
            }

            // n차 통합 코드
            String integratedCode = hCode + jibunNm + dongNm + hosuNm;

            System.out.println("고객명: "+ cliNm + " | n차 통합코드: " + integratedCode);
        }
    }

    
}
