package excel.style;

import org.apache.poi.ss.usermodel.*;

public class DefaultBodyStyle implements CellStyleStrategy {
    @Override
    public CellStyle getCellStyle(Workbook target) {
        CellStyle bodyStyle = target.createCellStyle();
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        bodyStyle.setBorderRight(BorderStyle.THIN);
        bodyStyle.setBorderLeft(BorderStyle.THIN);
        bodyStyle.setBorderTop(BorderStyle.THIN);
        bodyStyle.setBorderBottom(BorderStyle.THIN);
        bodyStyle.setWrapText(true); // 자동줄바꿈 기능

        return bodyStyle;
    }
}
