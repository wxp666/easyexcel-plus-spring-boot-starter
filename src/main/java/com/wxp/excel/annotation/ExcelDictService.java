package com.wxp.excel.annotation;

import java.util.Map;

/**
 * @author wxp
 * @date 2025/1/10
 * @apiNote
 */
public interface ExcelDictService {

    /**
     * 根据字典标识获取字典列表
     *
     * @param code 字典值
     * @return 字典列表(key:code,value:value)
     */
    Map<String, String> getValueByCode(String code);

}
