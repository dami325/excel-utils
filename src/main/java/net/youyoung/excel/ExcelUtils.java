package net.youyoung.excel;

import net.youyoung.excel.annotation.ExcelColumn;
import net.youyoung.excel.annotation.ExcelFieldInfo;
import net.youyoung.excel.annotation.ExcelTitle;
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
     * @param <T>                  DTO
     *                             사용코드 예시
     *                             response.setContentType(ExcelUtils.EXCEL_MIME_TYPE); // EXCEL_MIME_TYPE = application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     *                             response.setHeader("Content-Disposition", String.format("attachment;filename=%s_%s.xlsx", "CHAIN_LIST", LocalDate.now()));
     *                             ExcelUtils.download(list, clazz, response.getOutputStream(), LocaleContextHolder.getLocale());
     * @param list                 다운받을 DTO 목록
     * @param clazz                DTO 클래스 정보
     * @param responseOutputStream contentType 과 header 정보가 입력된 response.getOutputStream() 필요
     */
    public static <T> void download(List<T> list, Class<T> clazz, OutputStream responseOutputStream, String sheetTitle, Locale locale) {

        if (list == null || clazz == null || responseOutputStream == null)
            throw new IllegalArgumentException("list or clazz cannot be null");

        if (locale == null) locale = Locale.KOREAN;

        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            SXSSFSheet sheet = workbook.createSheet();
            DataFormat dataFormat = workbook.createDataFormat();
            int rowNo = 0, cellNo = 0;
            Map<String, ExcelFieldInfo> fieldInfoMap = excelColumnMetaData(clazz, workbook);
            Set<String> fieldNames = fieldInfoMap.keySet();
            int contentSize = list.size();

            workbook.setSheetName(0, sheetTitle);

            // @ExcelTitle
            if (clazz.isAnnotationPresent(ExcelTitle.class)) {

                ExcelTitle excelTitle = clazz.getAnnotation(ExcelTitle.class);

                if (excelTitle.useSheetTitle()) {
                    Row titleRow = sheet.createRow(rowNo++);
                    //cell 생성 및 설정
                    Cell titleCell = titleRow.createCell(cellNo);

                    titleCell.setCellValue(sheetTitle);
                    titleCell.setCellStyle(excelTitle.titleStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook));
                }

                if (excelTitle.useTotal()) {
                    Row totlaRow = sheet.createRow(rowNo++);
                    Cell totalCell = totlaRow.createCell(cellNo);
                    totalCell.setCellValue(isLocaleKorean(locale) ? "전체 : " + contentSize : "Total : " + contentSize);
                }
            }

            //header
            Row headerRow = sheet.createRow(rowNo++);

            for (String fieldName : fieldNames) {
                ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);

                Cell cell = headerRow.createCell(cellNo++);

                cell.setCellStyle(fieldInfo.headerStyleStrategy());

                setCellValue(cell, isLocaleKorean(locale) ? fieldInfo.header() : fieldInfo.headerEn(), fieldInfo.columnDefault());
            }

            if (contentSize == 0) {
                workbook.write(responseOutputStream);
                return;
            }


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

    private static boolean isLocaleKorean(Locale locale) {
        return locale.equals(Locale.KOREAN);
    }

    private static <T> Map<String, ExcelFieldInfo> excelColumnMetaData(Class<T> clazz, SXSSFWorkbook workbook) throws InstantiationException, IllegalAccessException, InvocationTargetException, NoSuchMethodException {
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
                                        excelColumn.headerStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook),
                                        excelColumn.bodyStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook),
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
        if (cellValue instanceof Number number) {
            cell.setCellValue(number.doubleValue());
        }
        else if (cellValue instanceof LocalDateTime localDateTime) {
            cell.setCellValue(localDateTime);
        }
        else if (cellValue instanceof LocalDate localDate) {
            cell.setCellValue(localDate);
        }
        else if (cellValue instanceof Date date) {
            cell.setCellValue(date);
        }
        else if (cellValue instanceof Boolean aBoolean) {
            cell.setCellValue(aBoolean);
        }
        else {
            cell.setCellValue(cellValue == null ? defaultValue : cellValue.toString());
        }

    }

}
