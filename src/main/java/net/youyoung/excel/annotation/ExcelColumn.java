package net.youyoung.excel.annotation;

import net.youyoung.excel.style.CellStyleStrategy;
import net.youyoung.excel.style.DefaultBodyStyle;
import net.youyoung.excel.style.DefaultHeaderStyle;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * header 다운로드 받을 필드 헤더 이름 미입력시 headerEn 값으로 대체
 *
 * headerEn (다국어 지원용 ) 다운로드 받을 필드 헤더 영문이름 미입력시 header 값으로 대체
 *
 * width Cell 가로폭
 *
 * headerStyle CellStyle 를 반환하는 CellStyleStrategy 의 구현체 Class 정보
 *
 * bodyStyle CellStyle 를 반환하는 CellStyleStrategy 의 구현체 Class 정보
 *
 * format 필드 데이터타입에 사용가능한 format 정보
 *    ex) double 타입 필드 는 format = "#,##0.00"
 *
 * columnDefault data 가 null 일 경우 보여줄 기본값
 *
 * 사용 예시 )
 *      T
 *      @ExcelColumn(header = "회사명")
 *      private String companyName;
 *
 */
@Target(FIELD)
@Retention(RUNTIME)
public @interface ExcelColumn {
    String header() default "";

    String headerEn() default "";

    int width() default 4096;

    Class<? extends CellStyleStrategy> headerStyle() default DefaultHeaderStyle.class;

    Class<? extends CellStyleStrategy> bodyStyle() default DefaultBodyStyle.class;

    String format() default "";

    String columnDefault() default "";

}
