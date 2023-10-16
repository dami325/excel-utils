package excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public interface CellStyleStrategy {

    CellStyle getCellStyle(Workbook target);

}
