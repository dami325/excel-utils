package excel.annotation;


import excel.style.CellStyleStrategy;

public record ExcelFieldInfo
        (
                String header,
                String headerEn,
                int width,
                Class<? extends CellStyleStrategy> headerStyleStrategy,
                Class<? extends CellStyleStrategy> bodyStyleStrategy,
                String format,
                String columnDefault
        ) {}
