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

    /**
     * 是否合并
     * @return 是否合并列
     */
    boolean isMerge() default false;

    /**
     * 要合并的列
     * @return 合并列数组
     */
    int[] mergeColumn() default {};

    /**
     * 表头行数
     * 合并时需计算，不指定默认-1 则取ExcelProperty注解value的长度
     */
    int headNumber() default -1;

}
