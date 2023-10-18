package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.*;

/**
 * 리플렉션을 사용하기 위한 INSTANCE 상수 필요
 */
public class DefaultHeaderStyle implements CellStyleStrategy{

    @Override
    public CellStyle getCellStyle(CellStyle target) {

        target.setBorderRight(BorderStyle.THIN);
        target.setBorderLeft(BorderStyle.THIN);
        target.setBorderTop(BorderStyle.THIN);
        target.setBorderBottom(BorderStyle.THIN);
        target.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        target.setAlignment(HorizontalAlignment.CENTER);
        target.setVerticalAlignment(VerticalAlignment.CENTER);

        return target;
    }
}
