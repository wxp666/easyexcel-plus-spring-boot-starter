package com.wxp.excel.handler;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.DefaultStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.wxp.excel.annotation.ResponseExcel;
import com.wxp.excel.strategy.AutoColumnWidthStrategy;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.MediaTypeFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote 写excel处理器
 */
public class ExcelWriteHandler {


    @SneakyThrows(IOException.class)
    public void write(List<?> returnValue, ResponseExcel responseExcel, HttpServletResponse response) {
        // 文件名
        String fileName = String.format("%s%s", URLEncoder.encode(responseExcel.name(), "UTF-8"), responseExcel.suffix().getValue())
                .replaceAll("\\+", "%20");
        // 根据实际的文件类型找到对应的 contentType
        String contentType = MediaTypeFactory.getMediaType(fileName)
                .map(MediaType::toString)
                .orElse("application/vnd.ms-excel");
        response.setContentType(contentType);
        response.setCharacterEncoding("utf-8");
        response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;filename*=utf-8''" + fileName);
        EasyExcelFactory.write(response.getOutputStream(), returnValue.get(0).getClass()).excelType(responseExcel.suffix()).sheet(responseExcel.sheetName())
                .registerWriteHandler(new AutoColumnWidthStrategy()).registerWriteHandler(getStyleStrategy())
                .doWrite(returnValue);
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
