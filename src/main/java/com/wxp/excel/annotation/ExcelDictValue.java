package com.wxp.excel.annotation;

import java.lang.annotation.*;

/**
 * @author wxp
 * @date 2025/1/10
 * @apiNote
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelDictValue {

    String value();

}
