package net.youyoung.excel.annotation;

import net.youyoung.excel.style.CellStyleStrategy;
import net.youyoung.excel.style.DefaultTitleStyle;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

@Target(TYPE)
@Retention(RUNTIME)
public @interface ExcelTitle {

    boolean useTotal() default true;

    Class<? extends CellStyleStrategy> titleStyle() default DefaultTitleStyle.class;

    boolean useSheetTitle() default true;
}
