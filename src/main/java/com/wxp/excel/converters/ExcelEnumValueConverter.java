package com.wxp.excel.converters;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.wxp.excel.annotation.ExcelEnumValue;
import com.wxp.excel.exception.ExcelPlusException;

import java.lang.reflect.Method;

/**
 * @author wxp
 * @date 2023/3/7
 * @apiNote
 */
public class ExcelEnumValueConverter implements Converter<Object> {

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
        ExcelEnumValue enumValue = contentProperty.getField().getAnnotation(ExcelEnumValue.class);
        Class<?> enumClass = enumValue.value();
        Object enumConstant = enumClass.getEnumConstants()[0];
        Method method;
        try {
            method = enumClass.getDeclaredMethod("getByCode",Object.class);
        } catch (NoSuchMethodException e) {
            throw new ExcelPlusException("enum convert fail,getByCode方法未找到",e);
        }
        Object invoke;
        try {
            invoke = method.invoke(enumConstant, value);
        } catch (Exception e) {
            throw new ExcelPlusException("enum convert fail,getByCode方法调用失败",e);
        }
        return new WriteCellData<>(invoke.toString());
    }
}
