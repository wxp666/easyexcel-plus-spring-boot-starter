package com.wxp.excel.annotation;

import com.alibaba.excel.support.ExcelTypeEnum;

import java.lang.annotation.*;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote 导出excle注解
 */
@Documented
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ResponseExcel {

    /**
     * 文件名称 不能为空
     * @return name
     */
    String name();

    /**
     * excel文件后缀 默认.xlsx
     * @return suffix
     */
    ExcelTypeEnum suffix() default ExcelTypeEnum.XLSX;

    /**
     * sheet名称
     * @return sheet
     */
    String sheetName() default "";


}
