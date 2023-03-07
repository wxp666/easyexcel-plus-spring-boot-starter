package com.wxp.excel.handler;

import com.alibaba.excel.EasyExcelFactory;
import com.wxp.excel.annotation.ResponseExcel;
import lombok.SneakyThrows;
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
        EasyExcelFactory.write(response.getOutputStream(), returnValue.get(0).getClass()).excelType(responseExcel.suffix()).sheet(responseExcel.sheetName()).doWrite(returnValue);
    }
}
