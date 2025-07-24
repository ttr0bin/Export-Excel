package likelion.excel.service;

import likelion.excel.entity.User;
import likelion.excel.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ByteArrayResource;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Service;

@Service
@RequiredArgsConstructor
public class ExcelService {

    private final UserRepository userRepository;

    public Resource excelDownload() {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("User");

        int rowNum = 0;
        // 헤더 행 생성
        Row headerRow = sheet.createRow(rowNum++);  // 0번째 행에 헤더 생성
        String[] headers = {"id", "email", "password", "username"};
        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
        }

        // 데이터 행 생성
// 정적 데이터 prac
// header : {"순번", "이름", "나이", "성별", "연락처"}
//        Object[][] data = {
//                {1, "한국인", 35, "남", "010-0000-0000"},
//                {2, "박원희", 11, "여", "010-1234-0000"},
//                {3, "이국한", 23, "여", "010-5678-0000"},
//                {4, "김명희", 27, "여", "010-9010-0000"},
//                {5, "김철민", 29, "남", "010-8888-0000"},
//        };
//        // Sheet 내에 데이터 행 구성
//        for (int i = 0; i < data.length; i++) {
//            Row row = sheet.createRow(i + 1);
//            for (int j = 0; j < data[i].length; j++) {
//                // 각 행의 셀 생성
//                Cell cell = row.createCell(j);
//
//                // 각 행의 값 입력
//                if (data[i][j] instanceof String) {
//                    cell.setCellValue((String) data[i][j]);
//                } // 문자 처리
//                if (data[i][j] instanceof Integer) {
//                    cell.setCellValue((Integer) data[i][j]);
//                } // 숫자 처리
//            }
//        }

        List<User> userList = userRepository.findAll();
        for (User user : userList) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(2).setCellValue(user.getEmail());
            row.createCell(3).setCellValue(user.getPassword());
            row.createCell(1).setCellValue(user.getUsername());
        }


        // 열 너비 자동 조정 (가장 넓은 셀에 맞게 너비 설정)
        for (int i = 0; i < headers.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // 파일 생성
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        // 메모리상 엑셀 바이트 배열 -> spring에서 응답 가능한 리소스 객체로 랩핑
        ByteArrayResource resource = new ByteArrayResource(outputStream.toByteArray());
        return resource;
    }
}
