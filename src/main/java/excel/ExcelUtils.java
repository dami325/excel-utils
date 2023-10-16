package excel;

import excel.annotation.ExcelColumn;
import excel.annotation.ExcelFieldInfo;
import excel.style.CellStyleStrategy;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;


/**
 * implementation 'org.apache.poi:poi-ooxml:5.2.3'
 *
 * response.setContentType(ExcelUtils.EXCEL_MIME_TYPE);
 */
public class ExcelUtils<T> {

    private static final Logger log = LogManager.getLogger(ExcelUtils.class);

    public static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private ExcelUtils() {}

    /**
     * 다국어 처리
     * 엑셀 다운로드 메서드
     *
     * @param list 다운받을 DTO 목록
     * @param clazz DTO 클래스 정보
     * @param responseOutputStream contentType 과 header 정보가 입력된 response.getOutputStream() 필요
     * @param <T> DTO
     *
     * 사용코드 예시
     *      response.setContentType(ExcelUtils.EXCEL_MIME_TYPE); // EXCEL_MIME_TYPE = application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     *      response.setHeader("Content-Disposition", String.format("attachment;filename=%s_%s.xlsx", "CHAIN_LIST", LocalDate.now()));
     *
     *      ExcelUtils.download(list, clazz, response.getOutputStream(), LocaleContextHolder.getLocale());
     */
    public static <T> void download(List<T> list, Class<T> clazz, OutputStream responseOutputStream, Locale locale) {

        if (locale == null) {
            log.info("locale is null");
            locale = Locale.KOREAN;
        }

        log.info("locale.getCountry = {} locale.getLanguage = {} ",locale.getCountry(), locale.getLanguage());

        if (list == null || clazz == null || responseOutputStream == null) throw new IllegalArgumentException("list or clazz cannot be null");

        log.info("Excel Download Start");
        try(SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            SXSSFSheet sheet = workbook.createSheet();

            int rowNo = 0, cellNo = 0;

            Map<String, ExcelFieldInfo> fieldInfoMap = getExcelColumnData(clazz);

            Set<String> excelColumnFieldNames = fieldInfoMap.keySet();

            Row headerRow = sheet.createRow(rowNo++);

            setHeaderData(locale, workbook, cellNo, fieldInfoMap, excelColumnFieldNames, headerRow);

            if (list.isEmpty()) return;

            setBodyData(list, workbook, rowNo, fieldInfoMap, sheet, excelColumnFieldNames);

            workbook.write(responseOutputStream);
            log.info("Excel Download Complete");

        } catch (IllegalAccessException | NoSuchMethodException | InstantiationException | InvocationTargetException |
                 IOException | NoSuchFieldException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                responseOutputStream.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 헤더 데이터 추가
     */
    private static void setHeaderData(Locale locale, SXSSFWorkbook workbook, int cellNo, Map<String, ExcelFieldInfo> fieldInfoMap, Set<String> excelColumnFieldNames, Row headerRow) throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {
        for (String fieldName : excelColumnFieldNames) {
            ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);

            Cell cell = headerRow.createCell(cellNo++);

            CellStyleStrategy headerStyleStrategy = fieldInfo.headerStyleStrategy().getDeclaredConstructor().newInstance();
            cell.setCellStyle(headerStyleStrategy.getCellStyle(workbook));

            if(locale.equals(Locale.KOREAN)){
                setCellValue(cell, fieldInfo.header(),fieldInfo.columnDefault());
            } else {
                setCellValue(cell,fieldInfo.headerEn(),fieldInfo.columnDefault());
            }
        }
    }

    /**
     * 바디 데이터 추가
     */
    private static <T> void setBodyData(List<T> list, SXSSFWorkbook workbook, int rowNo, Map<String, ExcelFieldInfo> fieldInfoMap, SXSSFSheet sheet, Set<String> fieldNames) throws NoSuchFieldException, IllegalAccessException, InstantiationException, InvocationTargetException, NoSuchMethodException {
        int cellNo;
        for (Object column : list) {
            cellNo = 0;
            Row cloumnRow = sheet.createRow(rowNo++);

            for (String fieldName : fieldNames) {
                ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);
                sheet.setColumnWidth(cellNo, fieldInfo.width());

                Field field = column.getClass().getDeclaredField(fieldName);
                field.setAccessible(true);

                Cell cell = cloumnRow.createCell(cellNo++);

                Object ob = field.get(column);
                setCellValue(cell, ob, fieldInfo.columnDefault());

                //body style
                CellStyleStrategy cellStyleStrategy = fieldInfo.bodyStyleStrategy().getDeclaredConstructor().newInstance();
                CellStyle cellStyle = cellStyleStrategy.getCellStyle(workbook);
                DataFormat dataFormat = workbook.createDataFormat();
                cellStyle.setDataFormat(dataFormat.getFormat(fieldInfo.format()));
                cell.setCellStyle(cellStyle);

            }

        }
    }


    /**
     * 값 타입 체크
     */
    private static void setCellValue(Cell cell, Object cellValue,String defaultValue) {
        if (cellValue instanceof Number) {
            Number numberValue = (Number) cellValue;
            cell.setCellValue(numberValue.doubleValue());
            return;
        }
        else if(cellValue instanceof LocalDateTime){
            LocalDateTime localDateTime = (LocalDateTime) cellValue;
            cell.setCellValue(localDateTime);
            return;
        }
        else if (cellValue instanceof LocalDate) {
            LocalDate localDate = (LocalDate) cellValue;
            cell.setCellValue(localDate);
            return;
        } else if (cellValue instanceof Date) {
            Date dateValue = (Date) cellValue;
            cell.setCellValue(dateValue);
            return;
        } else if (cellValue instanceof Boolean) {
            Boolean booleanValue = (Boolean) cellValue;
            cell.setCellValue(booleanValue);
            return;
        }
        cell.setCellValue(cellValue == null ? defaultValue : cellValue.toString());
    }

    /**
     * 필드 어노테이션 값 Map 에 저장
     *
     * key : @ExcelColumn 이 붙은 필드명
     *
     * Map<String, ExcelFieldInfo> 사용방법 )
     *
     * ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);
     * fieldInfo.header(); 로 저장된값 접근 가능
     */
    private static <T> Map<String, ExcelFieldInfo> getExcelColumnData(Class<T> clazz) {
        Map<String, ExcelFieldInfo> fieldInfoMap = new LinkedHashMap<>();

        for (Field field : clazz.getDeclaredFields()) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                fieldInfoMap.put(
                        field.getName(),
                        new ExcelFieldInfo(
                                excelColumn.header().equals("") ? excelColumn.headerEn() : excelColumn.header(),
                                excelColumn.headerEn().equals("") ? excelColumn.header() : excelColumn.headerEn(),
                                excelColumn.width(),
                                excelColumn.headerStyle(),
                                excelColumn.bodyStyle(),
                                excelColumn.format(),
                                excelColumn.columnDefault()
                        )
                );
            }
        }
        return fieldInfoMap;
    }

}
