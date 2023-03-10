package com.wxp.excel.annotation;

/**
 * @author wxp
 * @date 2023/3/7
 * @apiNote
 */
public interface ExcelEnum<T>{

    String getByCode(T code);
}
