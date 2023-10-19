package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class DefaultTitleStyle implements CellStyleStrategy{

    @Override
    public CellStyle getCellStyle(SXSSFWorkbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 15);
        cellStyle.setFont(font);

        return cellStyle;
    }
}
