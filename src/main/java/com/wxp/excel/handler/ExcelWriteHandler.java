package com.wxp.excel.handler;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.DefaultStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.wxp.excel.annotation.ExcelMergeColumn;
import com.wxp.excel.annotation.ResponseExcel;
import com.wxp.excel.exception.ExcelPlusException;
import com.wxp.excel.strategy.AutoColumnWidthStrategy;
import com.wxp.excel.strategy.ExcelMergeStrategy;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.MediaTypeFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote 写excel处理器
 */
public class ExcelWriteHandler {


    @SneakyThrows(IOException.class)
    public <T> void write(List<T> returnValue, ResponseExcel responseExcel, HttpServletResponse response) {
        // 文件名
        String fileName = String.format("%s%s", URLEncoder.encode(responseExcel.name(), "UTF-8"), responseExcel.suffix().getValue())
                .replaceAll("\\+", "%20");
        // 根据实际的文件类型找到对应的 contentType
        String contentType = MediaTypeFactory.getMediaType(fileName)
                .map(MediaType::toString)
                .orElse("application/vnd.ms-excel");
        response.setContentType(contentType);
        response.setCharacterEncoding("utf-8");
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename=" + fileName);
        ExcelWriterSheetBuilder excelWriterSheetBuilder = EasyExcelFactory.write(response.getOutputStream(), returnValue.get(0).getClass()).excelType(responseExcel.suffix()).sheet(responseExcel.sheetName())
                .registerWriteHandler(new AutoColumnWidthStrategy()).registerWriteHandler(getStyleStrategy());
        if (responseExcel.isMerge()) {
            if (responseExcel.mergeColumn().length <= 0) {
                throw new ExcelPlusException("excel合并时，合并的列mergeColumn属性不能为空");
            }
            excelWriterSheetBuilder.registerWriteHandler(getMergeStrategy(responseExcel, returnValue));
        }
        excelWriterSheetBuilder.doWrite(returnValue);
    }

    private <T> ExcelMergeStrategy getMergeStrategy(ResponseExcel responseExcel, List<T> returnValue) {
        Class<?> tempClass = returnValue.get(0).getClass();
        List<Field> tempFieldList = new ArrayList<>();
        while (tempClass != null) {
            Collections.addAll(tempFieldList, tempClass.getDeclaredFields());
            // Get the parent class and give it to yourself
            tempClass = tempClass.getSuperclass();
        }
        Field mergeField = tempFieldList.stream().filter(field -> field.getAnnotation(ExcelMergeColumn.class) != null).findFirst()
                .orElseThrow(() -> new ExcelPlusException("excel合并未找到ExcelMergeColumn注解"));
        Map<Object, List<T>> mergeMap = returnValue.stream().collect(Collectors.groupingBy(e -> {
            Object value = null;
            try {
                mergeField.setAccessible(true);
                value = mergeField.get(e);
            } catch (IllegalAccessException ex) {
                ex.printStackTrace();
            }
            return value;
        }));

        // 要合并的行，key：起始行，value：结束行
        Map<Integer, Integer> mapMerge = new HashMap<>();
        // 要合并的起始行 = head的行数
        int index = responseExcel.headNumber();
        // index == -1 没有指定起始行
        if (index == -1) {
            ExcelProperty excelProperty = Optional.of(mergeField.getAnnotation(ExcelProperty.class)).orElseThrow(() -> new ExcelPlusException("excel合并列未找到ExcelProperty注解"));
            index = excelProperty.value().length;
        }
        // 分组后数据的顺序会变化，清空数据 计算合并时从新添加
        returnValue.clear();
        for (List<T> vos : mergeMap.values()) {
            returnValue.addAll(vos);
            mapMerge.put(index, index + vos.size() - 1);
            index += vos.size();
        }
        return new ExcelMergeStrategy(mapMerge, responseExcel.mergeColumn());
    }

    private HorizontalCellStyleStrategy getStyleStrategy() {
        //内容样式
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        //垂直居中,水平居中
        contentWriteCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentWriteCellStyle.setBorderLeft(BorderStyle.THIN);
        contentWriteCellStyle.setBorderTop(BorderStyle.THIN);
        contentWriteCellStyle.setBorderRight(BorderStyle.THIN);
        contentWriteCellStyle.setBorderBottom(BorderStyle.THIN);
        //设置 自动换行
        contentWriteCellStyle.setWrapped(true);
        // 字体策略
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short) 12);
        contentWriteCellStyle.setWriteFont(contentWriteFont);
        // 使用默认样式策略，头样式使用easyexcel默认
        DefaultStyle defaultStyle = new DefaultStyle();
        defaultStyle.setContentWriteCellStyleList(ListUtils.newArrayList(contentWriteCellStyle));
        return defaultStyle;
    }
}
