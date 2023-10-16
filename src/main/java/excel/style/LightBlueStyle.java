package excel.style;

import org.apache.poi.ss.usermodel.*;

public class LightBlueStyle implements CellStyleStrategy {

    @Override
    public CellStyle getCellStyle(Workbook target) {
        CellStyle headerStyle = target.createCellStyle();
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = target.createFont();
        font.setFontName("맑은 고딕");
        font.setFontHeight((short) (9 * 20));
        headerStyle.setFont(font);

        return headerStyle;
    }

}