package net.youyoung.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;

public interface CellStyleStrategy {
    CellStyle getCellStyle(CellStyle target);

}
