package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.*;

public class DefaultBodyStyle implements CellStyleStrategy {

    @Override
    public CellStyle getCellStyle(CellStyle target) {

        target.setAlignment(HorizontalAlignment.CENTER);
        target.setVerticalAlignment(VerticalAlignment.CENTER);
        target.setBorderRight(BorderStyle.THIN);
        target.setBorderLeft(BorderStyle.THIN);
        target.setBorderTop(BorderStyle.THIN);
        target.setBorderBottom(BorderStyle.THIN);
        target.setWrapText(true); // 자동줄바꿈 기능

        return target;
    }
}
