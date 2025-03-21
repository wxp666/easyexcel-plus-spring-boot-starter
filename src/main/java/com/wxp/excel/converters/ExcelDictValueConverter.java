package com.wxp.excel.converters;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.ReadCellData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.alibaba.excel.util.StringUtils;
import com.wxp.excel.annotation.ExcelDictService;
import com.wxp.excel.annotation.ExcelDictValue;
import com.wxp.excel.exception.ExcelPlusException;
import com.wxp.excel.utils.SpringUtil;
import lombok.extern.slf4j.Slf4j;

import java.util.HashMap;
import java.util.Map;

/**
 * @author wxp
 * @date 2023/3/7
 * @apiNote
 */
@Slf4j
public class ExcelDictValueConverter implements Converter<Object> {

    @Override
    public Class<?> supportJavaTypeKey() {
        return Object.class;
    }

    @Override
    public CellDataTypeEnum supportExcelTypeKey() {
        return CellDataTypeEnum.STRING;
    }

    @Override
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        Map<String, String> dictMap = getDictMap(contentProperty);
        String result = dictMap.get(value.toString());
        if (StringUtils.isBlank(result)) {
            log.error("dict convert fail,未找到code对应的字典值,code:{},dictMap:{}", value, dictMap);
            result = value.toString();
        }
        return new WriteCellData<>(result);
    }

    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        String stringValue = cellData.getStringValue();
        Map<String, String> dictMap = getDictMap(contentProperty);
        //根据dictMap的value值获取对应的key值
        Class<?> type = contentProperty.getField().getType();
        for (Map.Entry<String, String> entry : dictMap.entrySet()) {
            if (entry.getValue().equals(stringValue)) {
                return covertStringToFieldType(entry.getKey(), type);
            }
        }
        log.error("dict convert fail,未找到value对应的字典值,value:{},dictMap:{}", stringValue, dictMap);
        return null;
    }

    private Object covertStringToFieldType(String value, Class<?> type) {
        if (type == String.class) {
            return value;
        } else if (type == Integer.class) {
            return Integer.valueOf(value);
        } else if (type == Long.class) {
            return Long.valueOf(value);
        } else if (type == Double.class) {
            return Double.valueOf(value);
        } else if (type == Float.class) {
            return Float.valueOf(value);
        } else if (type == Short.class) {
            return Short.valueOf(value);
        } else if (type == Byte.class) {
            return Byte.valueOf(value);
        } else if (type == Boolean.class) {
            return Boolean.valueOf(value);
        } else if (type == Character.class) {
            return value.charAt(0);
        } else {
            throw new ExcelPlusException("dict convert fail,不支持的类型");
        }
    }

    private Map<String, String> getDictMap(ExcelContentProperty contentProperty) {
        ExcelDictValue dictValue = contentProperty.getField().getAnnotation(ExcelDictValue.class);
        if (dictValue == null) {
            throw new ExcelPlusException("dict convert fail,未找到 ExcelDictValue 注解");
        }
        ExcelDictService service;
        try {
            service = SpringUtil.getBean(ExcelDictService.class);
        } catch (Exception e) {
            throw new ExcelPlusException("dict convert fail,未找到 ExcelDictService bean", e);
        }
        Map<String, String> dictMap = service.getValueByCode(dictValue.value());
        if (dictMap == null) {
            log.error("dict convert fail,未找到对应的字典，字典表示：{}", dictValue.value());
            return new HashMap<>();
        }
        return dictMap;
    }

}
