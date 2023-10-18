package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.*;

public class LightBlueStyle implements CellStyleStrategy {

    @Override
    public CellStyle getCellStyle(CellStyle target) {

        target.setBorderRight(BorderStyle.THIN);
        target.setBorderLeft(BorderStyle.THIN);
        target.setBorderTop(BorderStyle.THIN);
        target.setBorderBottom(BorderStyle.THIN);
        target.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        target.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        target.setAlignment(HorizontalAlignment.CENTER);
        target.setVerticalAlignment(VerticalAlignment.CENTER);

        return target;
    }

}
