package com.wxp.excel.converters;

import com.alibaba.excel.converters.Converter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.GlobalConfiguration;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.metadata.property.ExcelContentProperty;
import com.wxp.excel.annotation.ExcelEnumValue;

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
    public WriteCellData<?> convertToExcelData(Object value, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) throws Exception {
        ExcelEnumValue enumValue = contentProperty.getField().getAnnotation(ExcelEnumValue.class);
        Class<?> enumClass = enumValue.value();
        Object enumConstant = enumClass.getEnumConstants()[0];
        Method method = enumClass.getDeclaredMethod("getByCode",Object.class);
        Object invoke = method.invoke(enumConstant, value);
        return new WriteCellData<>(invoke.toString());
    }
}
