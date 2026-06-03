package com.wxp.excel.converters;

import com.wxp.excel.annotation.ExcelEnumValue;
import com.wxp.excel.exception.ExcelPlusException;
import org.apache.fesod.sheet.converters.Converter;
import org.apache.fesod.sheet.enums.CellDataTypeEnum;
import org.apache.fesod.sheet.metadata.GlobalConfiguration;
import org.apache.fesod.sheet.metadata.data.ReadCellData;
import org.apache.fesod.sheet.metadata.data.WriteCellData;
import org.apache.fesod.sheet.metadata.property.ExcelContentProperty;

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
        if (enumValue == null) {
            throw new ExcelPlusException("enum转换失败,未找到 ExcelEnumValue 注解");
        }
        Class<?> enumClass = enumValue.value();
        Object enumConstant = enumClass.getEnumConstants()[0];
        Method method;
        try {
            method = enumClass.getDeclaredMethod("getByCode", Object.class);
        } catch (NoSuchMethodException e) {
            throw new ExcelPlusException("enum convert fail,getByCode方法未找到", e);
        }
        Object invoke;
        try {
            invoke = method.invoke(enumConstant, value);
        } catch (Exception e) {
            throw new ExcelPlusException("enum convert fail,getByCode方法调用失败", e);
        }
        return new WriteCellData<>(invoke.toString());
    }

    @Override
    public Object convertToJavaData(ReadCellData<?> cellData, ExcelContentProperty contentProperty, GlobalConfiguration globalConfiguration) {
        String stringValue = cellData.getStringValue();
        ExcelEnumValue enumValue = contentProperty.getField().getAnnotation(ExcelEnumValue.class);
        Class<?> enumClass = enumValue.value();
        Object enumConstant = enumClass.getEnumConstants()[0];
        Method method;
        try {
            method = enumClass.getDeclaredMethod("getCode", String.class);
        } catch (NoSuchMethodException e) {
            try {
                method = enumClass.getDeclaredMethod("getCode", Object.class);
            } catch (NoSuchMethodException ex) {
                throw new ExcelPlusException("excel 导入 enum 转换失败,getCode方法未找到", e);
            }
        }
        Object invoke;
        try {
            invoke = method.invoke(enumConstant, stringValue);
        } catch (Exception e) {
            throw new ExcelPlusException("enum convert fail,getByCode方法调用失败", e);
        }
        return invoke;
    }

}
