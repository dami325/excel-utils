package net.youyoung.excel;

import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import net.youyoung.excel.annotation.ExcelColumn;
import net.youyoung.excel.annotation.ExcelFieldInfo;
import net.youyoung.excel.annotation.ExcelTitle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.context.i18n.LocaleContextHolder;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.lang.NonNull;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * @author Judalm park
 */
public class ExcelUtils<T> {

    public static final String EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    private ExcelUtils() {}

    /**
     * 다국어 처리
     * 엑셀 다운로드 메서드.
     *
     * @param list 다운받을 DTO 리스트
     * @param clazz 리스트 DTO 클래스 정보
     * @param downloadFileName 다운받을 파일명
     */
    public static <T> void download(@NonNull List<T> list, @NonNull Class<T> clazz, String downloadFileName) {

        parameterValidation(list, clazz);

        Locale locale = LocaleContextHolder.getLocale();

        try (SXSSFWorkbook workbook = new SXSSFWorkbook()) {
            SXSSFSheet sheet = workbook.createSheet();
            DataFormat dataFormat = workbook.createDataFormat();
            int rowNo = 0, cellNo = 0;
            Map<String, ExcelFieldInfo> fieldInfoMap = excelColumnMetaData(clazz, workbook);
            Set<String> fieldNames = fieldInfoMap.keySet();
            int contentSize = list.size();

            String sheetTitle = "";
            // @ExcelTitle
            if (clazz.isAnnotationPresent(ExcelTitle.class)) {

                ExcelTitle excelTitle = clazz.getAnnotation(ExcelTitle.class);

                if (excelTitle.useSheetTitle()) {
                    Row titleRow = sheet.createRow(rowNo++);
                    //cell 생성 및 설정
                    Cell titleCell = titleRow.createCell(cellNo);

                    sheetTitle = isLocaleKorean(locale) ? excelTitle.sheetTitle() : excelTitle.sheetTitleEn();

                    if (sheetTitle.equals(""))
                        sheetTitle = downloadFileName;


                    titleCell.setCellValue(sheetTitle);
                    titleCell.setCellStyle(excelTitle.titleStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook));
                }

                if (excelTitle.useTotal()) {
                    Row totlaRow = sheet.createRow(rowNo++);
                    Cell totalCell = totlaRow.createCell(cellNo);
                    totalCell.setCellValue(isLocaleKorean(locale) ? "전체 : " + contentSize : "Total : " + contentSize);
                }
            }

            setSheetTitle(workbook, sheetTitle);

            //header
            rowNo = setHeaderCellValue(locale, sheet, rowNo, cellNo, fieldInfoMap, fieldNames);

            if (contentSize == 0) {
                write(downloadFileName, workbook);
                return;
            }


            //body
            setBodyCellValue(list, sheet, dataFormat, rowNo, fieldInfoMap, fieldNames);

            write(downloadFileName, workbook);

        } catch (IllegalAccessException | IOException | NoSuchFieldException | NoSuchMethodException |
                 InstantiationException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * @return 엑셀 리소스만 필요할때
     * 후처리는 사용자에게 위임
     * 사용코드
     *
     * resource = ExcelUtils.getResource(excelList, clazz);
     * HttpHeaders headers = new HttpHeaders();
     * headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + orgFileName + "\"");
     * headers.setContentType(MediaType.parseMediaType(mimeType));
     * new ResponseEntity<T>(resource, httpHeaders, HttpStatus.OK);
     */
    public static <T> Resource getResource(@NonNull List<T> list, @NonNull Class<T> clazz) {

        parameterValidation(list, clazz);

        Locale locale = LocaleContextHolder.getLocale();

        try (SXSSFWorkbook workbook = new SXSSFWorkbook();
             ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();) {
            SXSSFSheet sheet = workbook.createSheet();
            DataFormat dataFormat = workbook.createDataFormat();
            int rowNo = 0, cellNo = 0;
            Map<String, ExcelFieldInfo> fieldInfoMap = excelColumnMetaData(clazz, workbook);
            Set<String> fieldNames = fieldInfoMap.keySet();
            int contentSize = list.size();

            String sheetTitle = "";
            // @ExcelTitle
            if (clazz.isAnnotationPresent(ExcelTitle.class)) {

                ExcelTitle excelTitle = clazz.getAnnotation(ExcelTitle.class);

                if (excelTitle.useSheetTitle()) {
                    Row titleRow = sheet.createRow(rowNo++);
                    //cell 생성 및 설정
                    Cell titleCell = titleRow.createCell(cellNo);

                    sheetTitle = isLocaleKorean(locale) ? excelTitle.sheetTitle() : excelTitle.sheetTitleEn();

                    titleCell.setCellValue(sheetTitle);
                    titleCell.setCellStyle(excelTitle.titleStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook));
                }

                if (excelTitle.useTotal()) {
                    Row totlaRow = sheet.createRow(rowNo++);
                    Cell totalCell = totlaRow.createCell(cellNo);
                    totalCell.setCellValue(isLocaleKorean(locale) ? "전체 : " + contentSize : "Total : " + contentSize);
                }
            }

            setSheetTitle(workbook, sheetTitle);

            //header
            rowNo = setHeaderCellValue(locale, sheet, rowNo, cellNo, fieldInfoMap, fieldNames);

            if (contentSize == 0)
                return getWorkbookResource(workbook, byteArrayOutputStream);


            //body
            setBodyCellValue(list, sheet, dataFormat, rowNo, fieldInfoMap, fieldNames);

            return getWorkbookResource(workbook, byteArrayOutputStream);

        } catch (IllegalAccessException | IOException | NoSuchFieldException | NoSuchMethodException |
                 InstantiationException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 시트 첫번째 셀 제목 커스텀 추가
     */
    public static <T> Resource getResource(@NonNull List<T> list, @NonNull Class<T> clazz, String titleAppend) {

        parameterValidation(list, clazz);

        Locale locale = LocaleContextHolder.getLocale();

        try (SXSSFWorkbook workbook = new SXSSFWorkbook();
             ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();) {
            SXSSFSheet sheet = workbook.createSheet();
            DataFormat dataFormat = workbook.createDataFormat();
            int rowNo = 0, cellNo = 0;
            Map<String, ExcelFieldInfo> fieldInfoMap = excelColumnMetaData(clazz, workbook);
            Set<String> fieldNames = fieldInfoMap.keySet();
            int contentSize = list.size();

            String sheetTitle = "";
            // @ExcelTitle
            if (clazz.isAnnotationPresent(ExcelTitle.class)) {

                ExcelTitle excelTitle = clazz.getAnnotation(ExcelTitle.class);

                if (excelTitle.useSheetTitle()) {
                    Row titleRow = sheet.createRow(rowNo++);
                    //cell 생성 및 설정
                    Cell titleCell = titleRow.createCell(cellNo);

                    sheetTitle = isLocaleKorean(locale) ? excelTitle.sheetTitle() : excelTitle.sheetTitleEn();

                    sheetTitle = sheetTitle.concat(titleAppend);

                    titleCell.setCellValue(sheetTitle);
                    titleCell.setCellStyle(excelTitle.titleStyle().getDeclaredConstructor().newInstance().getCellStyle(workbook));
                }

                if (excelTitle.useTotal()) {
                    Row totlaRow = sheet.createRow(rowNo++);
                    Cell totalCell = totlaRow.createCell(cellNo);
                    totalCell.setCellValue(isLocaleKorean(locale) ? "전체 : " + contentSize : "Total : " + contentSize);
                }
            }

            setSheetTitle(workbook, sheetTitle);

            //header
            rowNo = setHeaderCellValue(locale, sheet, rowNo, cellNo, fieldInfoMap, fieldNames);

            if (contentSize == 0)
                return getWorkbookResource(workbook, byteArrayOutputStream);

            //body
            setBodyCellValue(list, sheet, dataFormat, rowNo, fieldInfoMap, fieldNames);

            return getWorkbookResource(workbook, byteArrayOutputStream);

        } catch (IllegalAccessException | IOException | NoSuchFieldException | NoSuchMethodException |
                 InstantiationException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * cell 헤더 값 입력
     */
    private static int setHeaderCellValue(Locale locale, SXSSFSheet sheet, int rowNo, int cellNo, Map<String, ExcelFieldInfo> fieldInfoMap, Set<String> fieldNames) {
        Row headerRow = sheet.createRow(rowNo++);

        for (String fieldName : fieldNames) {
            ExcelFieldInfo fieldInfo = fieldInfoMap.get(fieldName);

            Cell cell = headerRow.createCell(cellNo++);

            cell.setCellStyle(fieldInfo.headerStyleStrategy());

            setCellValue(cell, isLocaleKorean(locale) ? fieldInfo.header() : fieldInfo.headerEn(), fieldInfo.columnDefault());
        }
        return rowNo;
    }

    /**
     * Body 리스트 값 입력
     */
    private static <T> void setBodyCellValue(List<T> list, SXSSFSheet sheet, DataFormat dataFormat, int rowNo, Map<String, ExcelFieldInfo> fieldInfoMap, Set<String> fieldNames) throws NoSuchFieldException, IllegalAccessException {
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

                //body value
                Object ob = field.get(column);
                setCellValue(cell, ob, fieldInfo.columnDefault());

                //body style
                CellStyle cellStyle = fieldInfo.bodyStyleStrategy();
                cellStyle.setDataFormat(dataFormat.getFormat(fieldInfo.format()));
                cell.setCellStyle(cellStyle);

            }

        }
    }

    /***
     * 시트명 미입력 시 기본값 추가
     */
    private static void setSheetTitle(SXSSFWorkbook workbook, String sheetTitle) {
        workbook.setSheetName(0, sheetTitle.equals("") ? "Sheet1" : sheetTitle);
    }

    private static <T> void parameterValidation(List<T> list, Class<T> clazz) {
        if (list == null || clazz == null)
            throw new IllegalArgumentException("list or clazz cannot be null");
    }

    /**
     * 리소스 반환
     */
    private static Resource getWorkbookResource(SXSSFWorkbook workbook, ByteArrayOutputStream byteArrayOutputStream) throws IOException {
        workbook.write(byteArrayOutputStream);
        byte[] excelFileBytes = byteArrayOutputStream.toByteArray();
        Resource resource = new ByteArrayResource(excelFileBytes);
        workbook.close();
        workbook.dispose();
        return resource;
    }

    /**
     * 엑셀 다운로드 실행
     */
    private static void write(String downloadFileName, SXSSFWorkbook workbook){
        HttpServletResponse response = getHttpServletResponse();
        try(ServletOutputStream outputStream = response.getOutputStream()){
            response.setContentType(EXCEL_MIME_TYPE);
            String encodedFileName = encodeFileName(String.format("%s_%s.xlsx", downloadFileName, LocalDate.now()));
            response.setHeader("Content-Disposition", "attachment;filename=" + encodedFileName);

            workbook.write(outputStream);
            workbook.close();
            workbook.dispose();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     *파일명을 안전하게 인코딩
     */
    private static String encodeFileName(String fileName) throws UnsupportedEncodingException {
        return URLEncoder.encode(fileName, StandardCharsets.UTF_8.toString())
                .replaceAll("\\+", "%20")
                .replaceAll("%21", "!")
                .replaceAll("%27", "'")
                .replaceAll("%28", "(")
                .replaceAll("%29", ")")
                .replaceAll("%7E", "~");
    }

    /**
     * 요청 리스폰스 얻어오기
     */
    private static HttpServletResponse getHttpServletResponse() {
        ServletRequestAttributes requestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = requestAttributes.getResponse();
        return response;
    }

    private static boolean isLocaleKorean(Locale locale) {
        return locale.equals(Locale.KOREAN);
    }

    /**
     * 레코더에 @ExcelColumn 어노테이션 정보 저장
     */
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
     * 값 타입 체크후 셀 형변환
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
