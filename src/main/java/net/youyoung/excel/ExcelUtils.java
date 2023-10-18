package net.youyoung.excel;

import net.youyoung.excel.annotation.ExcelColumn;
import net.youyoung.excel.annotation.ExcelFieldInfo;
import net.youyoung.excel.style.CellStyleStrategy;
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
 * <p>
 * response.setContentType(ExcelUtils.EXCEL_MIME_TYPE);
 */
public class ExcelUtils<T> {

    public static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private ExcelUtils() {
    }

    /**
     * 다국어 처리
     * 엑셀 다운로드 메서드
     *
     * @param list                 다운받을 DTO 목록
     * @param clazz                DTO 클래스 정보
     * @param responseOutputStream contentType 과 header 정보가 입력된 response.getOutputStream() 필요
     * @param <T>                  DTO
     *                             <p>
     *                             사용코드 예시
     *                             response.setContentType(ExcelUtils.EXCEL_MIME_TYPE); // EXCEL_MIME_TYPE = application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     *                             response.setHeader("Content-Disposition", String.format("attachment;filename=%s_%s.xlsx", "CHAIN_LIST", LocalDate.now()));
     *                             <p>
     *                             ExcelUtils.download(list, clazz, response.getOutputStream(), LocaleContextHolder.getLocale());
     */
    public static <T> void download(List<T> list, Class<T> clazz, OutputStream responseOutputStream, Locale locale) {

        if (locale == null) {
            locale = Locale.KOREAN;
        }

        if (list == null || clazz == null || responseOutputStream == null)
            throw new IllegalArgumentException("list or clazz cannot be null");

        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            SXSSFSheet sheet = workbook.createSheet();
            DataFormat dataFormat = workbook.createDataFormat();
            int rowNo = 0, cellNo = 0;
            Map<String, ExcelFieldInfo> fieldInfoMap = enumColumnMetaData(clazz, workbook);
            Set<String> fieldNames = fieldInfoMap.keySet();

            //header
            Row headerRow = sheet.createRow(rowNo++);

            for (String fieldName : fieldNames) {
                ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);

                Cell cell = headerRow.createCell(cellNo++);

                cell.setCellStyle(fieldInfo.headerStyleStrategy());

                if (locale.equals(Locale.KOREAN)) {
                    setCellValue(cell, fieldInfo.header(), fieldInfo.columnDefault());
                }
                else {
                    setCellValue(cell, fieldInfo.headerEn(), fieldInfo.columnDefault());
                }
            }

            if (list.isEmpty()) return;


            //body
            for (Object column : list) {
                cellNo = 0;
                Row cloumnRow = sheet.createRow(rowNo++);

                for (String fieldName : fieldNames) {
                    ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);
                    sheet.setColumnWidth(cellNo, fieldInfo.width());

                    Field field = column.getClass().getDeclaredField(fieldName);
                    field.setAccessible(true);

                    Cell cell = cloumnRow.createCell(cellNo++);

                    //body value
                    Object ob = field.get(column);
                    setCellValue(cell, ob, fieldInfo.columnDefault());

                    //body style
                    CellStyle cellStyle = fieldInfo.bodyStyleStrategy();
                    cellStyle.setDataFormat(dataFormat.getFormat(fieldInfo.format()));
                    cell.setCellStyle(cellStyle);

                }

            }

            workbook.write(responseOutputStream);

        } catch (IllegalAccessException | IOException | NoSuchFieldException | NoSuchMethodException |
                 InstantiationException | InvocationTargetException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                responseOutputStream.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static <T> Map<String, ExcelFieldInfo> enumColumnMetaData(Class<T> clazz, SXSSFWorkbook workbook) throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {
        Map<String, ExcelFieldInfo> fieldInfoMap = new LinkedHashMap<>();

        for (Field field : clazz.getDeclaredFields()) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {

                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

                fieldInfoMap.put(
                        field.getName(),
                        new ExcelFieldInfo
                                (
                                        excelColumn.header().equals("") ? excelColumn.headerEn() : excelColumn.header(),
                                        excelColumn.headerEn().equals("") ? excelColumn.header() : excelColumn.headerEn(),
                                        excelColumn.width(),
                                        excelColumn.headerStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook.createCellStyle()),
                                        excelColumn.bodyStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook.createCellStyle()),
                                        excelColumn.format(),
                                        excelColumn.columnDefault()
                                )
                );
            }
        }
        return fieldInfoMap;
    }


    /**
     * 값 타입 체크
     */
    private static void setCellValue(Cell cell, Object cellValue, String defaultValue) {
        if (cellValue instanceof Number) {
            Number numberValue = (Number) cellValue;
            cell.setCellValue(numberValue.doubleValue());
            return;
        } else if (cellValue instanceof LocalDateTime) {
            LocalDateTime localDateTime = (LocalDateTime) cellValue;
            cell.setCellValue(localDateTime);
            return;
        } else if (cellValue instanceof LocalDate) {
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

}
