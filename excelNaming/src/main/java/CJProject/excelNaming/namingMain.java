package CJProject.excelNaming;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class namingMain {
    // 로그 패키지
    private static final Logger logger = LogManager.getLogger(namingMain.class);

    public static void main(String[] args) throws IOException {
        // 설정 파일 가져오기
        Properties properties = new Properties();
        String configFile = "config.properties";

        String logFilePath = Paths.get("logs", "app.log").toAbsolutePath().toString();
        System.out.println("Log file path: " + logFilePath);

     // UTF-8로 파일 읽기
        try (InputStream input = new FileInputStream(configFile);
             InputStreamReader reader = new InputStreamReader(input, StandardCharsets.UTF_8)) {
            properties.load(reader);
        }

        // 설정파일에서 값 가져오기
        // 파일 경로
        String directoryPath = properties.getProperty("directoryPath");

        // 제목으로 쓸 셀 리스트 수, 리스트
        int cellListCount = Integer.parseInt(properties.getProperty("cellListCount"));

        List<List<String>> cellList = new ArrayList<>();
        List<String> fileIdentifyCellList = new ArrayList<>();
        List<String> identifyCharList = new ArrayList<>();
        List<Integer> dateCellList = new ArrayList<>();
        List<List<Integer>> dateParsingList = new ArrayList<>();

        if (cellListCount != 0) {
            try {
                for (int i = 0; i < cellListCount; i++) {
                    String cellvar = properties.getProperty("cellList" + (i + 1));
                    cellList.add(Arrays.asList(cellvar.split(",")));
                }
            } catch (Exception e) {
                logger.info("cellListCount 값이 없거나 cellList가 없습니다. 종료합니다.");
                System.out.println("cellListCount 값이 없거나 cellList가 없습니다. 종료합니다.");
                System.exit(1);
            }

            try {
                for (int i = 0; i < cellListCount; i++) {
                    String identifychar = properties.getProperty("fileIdentifyCell" + (i + 1));
                    fileIdentifyCellList.add(identifychar);
                }
            } catch (Exception e) {
                logger.info("fileIdentifyCellList가 없습니다. 종료합니다.");
                System.out.println("fileIdentifyCellList가 없습니다. 종료합니다.");
                System.exit(1);
            }

            try {
                for (int i = 0; i < cellListCount; i++) {
                    String charVar = properties.getProperty("IdentifyChar" + (i + 1));
                    identifyCharList.add(charVar);
                }
            } catch (Exception e) {
                logger.info("IdentifyCharList가 없습니다. 종료합니다.");
                System.out.println("IdentifyCharList가 없습니다. 종료합니다.");
                System.exit(1);
            }
            try {
                for (int i = 0; i < cellListCount; i++) {
                    int dateCellvar = Integer.parseInt(properties.getProperty("dateCell" + (i + 1)));
                    dateCellList.add(dateCellvar);
                }
            } catch (Exception e) {
                logger.info("dateCell가 없습니다. 종료합니다.");
                System.out.println("dateCell가 없습니다. 종료합니다.");
                System.exit(1);
            }
            try {
                for (int i = 0; i < cellListCount; i++) {
                	String dateParsingVar = properties.getProperty("dateParsing" + (i + 1));
                	String[] parts = dateParsingVar.split(",");
                	List<Integer> intVar = new ArrayList<>();
                    for (String part : parts) {
                    	intVar.add(Integer.parseInt(part.trim()));
                    }
                    dateParsingList.add(intVar);
                }
            } catch (Exception e) {
                logger.info("dateParsing가 없습니다. 종료합니다.");
                System.out.println("dateParsing가 없습니다. 종료합니다.");
                System.exit(1);
            }
        }

        logger.info("directoryPath: " + directoryPath);
        System.out.println("directoryPath: " + directoryPath);

        // 디렉토리 생성
        File file = new File(directoryPath + "\\copyFiles");
        boolean directoryCreated = file.mkdir();
        if (directoryCreated)
            logger.info("copyFiles폴더 생성");

        // 파일 식별 및 처리
        try {
            Files.list(Paths.get(directoryPath))
                .filter(Files::isRegularFile)
                .filter(path -> path.toString().endsWith(".tmp"))
                .forEach(path -> {
                    try {
                        processFile(path, directoryPath, cellListCount, cellList, 
                        		fileIdentifyCellList, identifyCharList, dateCellList, dateParsingList);
                    } catch (IOException e) {
                        logger.info("================processFile 도중 에러 발생================");
                        System.out.println("================processFile 도중 에러 발생================");
                        e.printStackTrace();
                    }
                });
        } catch (IOException e) {
            logger.info("================Files.list 도중 에러 발생================");
            System.out.println("================Files.list 도중 에러 발생================");
            e.printStackTrace();
        }
    }

    private static void processFile(Path path, String directoryPath,
            int cellListCount, List<List<String>> cellList,
            List<String> fileIdentifyCellList, List<String> identifyCharList,
            List<Integer> dateCellList, List<List<Integer>> dateParsingList) throws IOException {
        try (FileInputStream fis = new FileInputStream(path.toFile()); Workbook workbook = new XSSFWorkbook(fis)) {
            logger.info("================================");
            logger.info("path : " + path);
            System.out.println("================================");
            System.out.println("path : " + path);

            Sheet sheet = workbook.getSheetAt(0);
//            결과 문자열
            String rs = "";
            
//          파일 식별하는 코드로 어떤 파일인지 식별

            int cnt = 0;
            for(int i=0;i<cellListCount;i++) {
            	if(cellInString(sheet, fileIdentifyCellList.get(i), cnt).startsWith(identifyCharList.get(i))) {
            		cnt=i;
            		break;
            	}
            }
            List<String> cellSelectList = cellList.get(cnt);
            for (int i=0;i<cellSelectList.size(); i++) {
            	String cellCode = cellSelectList.get(i);
                String cellString = cellInString(sheet, cellCode, cnt);
                if(dateCellList.get(cnt) != 0 && (dateCellList.get(cnt)-1) == i) {
                	cellString = dateParsing(cellString, dateParsingList.get(cnt));
                }
                logger.info("Cell Value, " + cellCode + " : " + cellString);
                System.out.println("Cell Value, " + cellCode + " : " + cellString);
                rs = rs + cellString + "_";
            }
            logger.info(rs);
            System.out.println(rs);

            Path destinationPath = Paths.get(directoryPath, "copyFiles", rs + ".xlsx");
            Files.copy(path, destinationPath);
            logger.info("================================");
            System.out.println("================================");
        } catch (Exception e) {
            e.printStackTrace();
            logger.info("변환하던 중 알 수 없는 이유로 종료합니다. 자세한 사항은 시스템 로그를 확인하십시오.");
            System.out.println("변환하던 중 알 수 없는 이유로 종료합니다. 자세한 사항은 시스템 로그를 확인하십시오.");
            System.exit(1);
        }
    }

    private static String cellInString(Sheet sheet, String cellCode, int cnt) {
        if ("AA".equals(cellCode)) {
            return "출금전표";
        }
        CellReference cellRef = new CellReference(cellCode);
        Row row = sheet.getRow(cellRef.getRow());
        Cell cell = row.getCell(cellRef.getCol());
        String cellString = "";

        if (cell != null) {
            if (cell.getCellType() == CellType.STRING) {
                cellString = cell.getStringCellValue();
//                // 출금전표일때 월 계산
//                if (cellCode.equals("C9")) {
//                    cellString = cellString.substring(4, 7) + "월";
//                }
//                // 품의서일때 월 계산
//                if (cellCode.equals("I4")) {
//                    cellString = cellString.substring(cellString.length() - 2) + "월";
//                }
//                if (cellString.endsWith("월월")) {
//                    cellString = cellString.substring(0, cellString.length() - 1);
//                }
//                cellString = cellString.replaceAll(" ", "");
            } else {
                logger.info("Cell is not a string or is empty.");
                System.out.println("Cell is not a string or is empty.");
            }
        } else {
            logger.info("Cell is empty or does not exist.");
            System.out.println("Cell is empty or does not exist.");
        }
        return cellString;
    }
    private static String dateParsing(String dateString, List<Integer> dateParsingList) {
    	try {
    		dateString = dateString.substring(dateParsingList.get(0)-1, dateParsingList.get(1));
    	} catch (Exception e) {
    		dateString = dateString.substring(dateParsingList.get(0)-1, dateString.length() - 1);
    		logger.info("날짜 위치를 다시 설정해주세요.");
    		e.printStackTrace();
    	}
    	dateString = dateString.replaceAll("[^0-9]", "");			
    	if(!dateString.endsWith("월")){
    		dateString = dateString  + "월";    		
    	}
    	System.out.println(dateString);
    	return dateString;
    }
}
