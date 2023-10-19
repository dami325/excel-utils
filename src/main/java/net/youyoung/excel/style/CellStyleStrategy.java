package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

@FunctionalInterface
public interface CellStyleStrategy {
    CellStyle getCellStyle(SXSSFWorkbook workbook);

}

