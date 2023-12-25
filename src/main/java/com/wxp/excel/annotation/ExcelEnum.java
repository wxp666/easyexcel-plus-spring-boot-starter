package com.wxp.excel.annotation;

/**
 * @author wxp
 * @date 2023/3/7
 * @apiNote
 */
public interface ExcelEnum<T> {

    /**
     * 根据字典值获取描述
     *
     * @param code 字典值
     * @return 描述
     */
    String getByCode(T code);

    /**
     * 根据描述获取字典值
     *
     * @param name 描述
     * @return 字典值
     */
    T getCode(String name);
}
