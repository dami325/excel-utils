package net.youyoung.excel.annotation;


import org.apache.poi.ss.usermodel.CellStyle;

public record ExcelFieldInfo
        (
                String header,
                String headerEn,
                int width,
                CellStyle headerStyleStrategy,
                CellStyle bodyStyleStrategy,
                String format,
                String columnDefault
        ) {}
