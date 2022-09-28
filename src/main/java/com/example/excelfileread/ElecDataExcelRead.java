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
public class ElecDataExcelRead {
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

        ReadOption elecExcel = new ReadOption();
//        ro.setFilePath("/Users/taehyeonkim/Desktop/incuvers/통합고지 가스전기데이터/가스 계약자 DB.xlsx");
        elecExcel.setFilePath("/Users/taehyeonkim/Desktop/electrical.xlsx");
        elecExcel.setOutputColumns("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P");
        elecExcel.setStartRow(3);


        List<Map<String,String>> result = GasDataExcelRead.read(elecExcel);

        // 행안부 주소 검색 Api 로 뽑아온 행정동코드 저장할 변수
        String hCode = "";

        for(Map<String,String> map : result){
            // 사용자 information
            String elecCliNm = map.get("A");            // 고객명
            String conCode = map.get("B");              // 계약코드
            String cliCode = map.get("C");              // 고객코드
            String elecPayNm = map.get("D");            // 전자납부번호

            // 소유자 주소
            String ownPostNm = map.get("E");            // 우편번호
            String ownAddress = map.get("F");           // 주소
            String ownDetailAdd = map.get("G");         // 상세주소
            String ownDoroAdd = map.get("H");           // 도로명주소

            // 전기사용 장소
            String usePostNm = map.get("I");            // 우편번호
            String useAddress = map.get("J");           // 주소
            String useDetailAdd = map.get("K");         // 상세주소
            String useDoroAdd = map.get("L");           // 도로명주소

            // 요금청구 주소
            String billPpostNm = map.get("M");          // 우편번호
            String billAddress = map.get("N");          // 주소
            String billDetailAdd = map.get("O");        // 상세주소
            String billDoroAdd = map.get("P");          // 도로명주소

            try {

                // 주소 검색, URLEncoder는 URL을 인코딩 하기위해 사용하는 클래스
                String keyword = URLEncoder.encode(useAddress, "UTF-8");

                URL url = new URL("https://business.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=10&keyword="
                        + keyword + "&confmKey=devU01TX0FVVEgyMDIyMDkwMjE2MTYzOTExMjk0NDM=&resultType=json");

                BufferedReader bufferedReader;


                bufferedReader = new BufferedReader(new InputStreamReader(url.openStream(), "UTF-8"));

                String res = bufferedReader.readLine();

                // String 값을 JSON 형태로 추출하기 위해 사용하는 라이브러리
                JSONParser jsonParser = new JSONParser();
                JSONObject jsonObject = (JSONObject)jsonParser.parse(res);
                JSONObject addResult = (JSONObject)jsonObject.get("results");

                // 리스트 추출
                JSONArray jusoArray = (JSONArray)addResult.get("juso");


                // IndexOutOfBoundsException: index:0, Size: 0 관련 에러 발생으로 인한 조건문
                if (jusoArray.size()!=0){

                    // 컬렉션 추출 주소정보 뽑을 준비 완료!
                    JSONObject jusoColl = (JSONObject) jusoArray.get(0);

                    hCode = jusoColl.get("admCd").toString();              // 행정동코드 10자리
                }

            }

            catch (Exception e) {
                log.error("잘못된 접근입니다",e);
            }

//            // 숫자 빼고 모두 제거
            String match = "[^0-9]";
            useAddress = useAddress.replaceAll(match, "");
            useDetailAdd = useDetailAdd.replaceAll(match, "");

//
//            // 건물 동의 값이 null 인 경우
            if(useDetailAdd.equals("")){
                useDetailAdd = "0";
            }

//
//            // int 로 형변환 이유는 동과 호수 변수 앞에 0 정수로 채우기 위한 작업(통합코드 생성시 규칙)
            int intJibun = Integer.parseInt(useAddress);
            int intDongHosu = Integer.parseInt(useDetailAdd);

//
            String jibunNm = String.format("%06d",intJibun);           // 최종 지번 숫자만(6자리)
            String dongHosu = String.format("%08d",intDongHosu);       // 최종 동호수 숫자만(8자리)


            // n차 통합 코드
            StringBuilder integratedCode = new StringBuilder();
            integratedCode.append(hCode);
            integratedCode.append(jibunNm);
            integratedCode.append(dongHosu);


            log.info("고객명: "+ elecCliNm + " | n차 통합코드: " + integratedCode);
        }
    }
}
