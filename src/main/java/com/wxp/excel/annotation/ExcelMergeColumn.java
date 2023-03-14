package com.wxp.excel.annotation;

import java.lang.annotation.*;

/**
 * @author wxp
 * @date 2023/3/13
 * @apiNote
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelMergeColumn {
}
