package com.wxp.excel.handler;

import com.wxp.excel.annotation.ResponseExcel;
import lombok.RequiredArgsConstructor;
import org.springframework.core.MethodParameter;
import org.springframework.util.Assert;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.method.support.ModelAndViewContainer;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote excel返回值处理器
 */
@RequiredArgsConstructor
public class ExcelReturnValueHandler implements HandlerMethodReturnValueHandler {

    private final ExcelWriteHandler excelWriteHandler;

    @Override
    public boolean supportsReturnType(MethodParameter returnType) {
        // 判断方法是否标注ResponseExcel注解
        return returnType.getMethodAnnotation(ResponseExcel.class) != null;
    }

    @Override
    public void handleReturnValue(Object returnValue, MethodParameter returnType, ModelAndViewContainer mavContainer, NativeWebRequest webRequest) throws Exception {
        HttpServletResponse response = webRequest.getNativeResponse(HttpServletResponse.class);
        Assert.notNull(response,"Excel导出HttpServletResponse为空");
        ResponseExcel responseExcel = returnType.getMethodAnnotation(ResponseExcel.class);
        Assert.notNull(responseExcel,"Excel导出ResponseExcel为空");
        mavContainer.setRequestHandled(true);
        // 判断返回值是否是list类型
        if ((returnValue instanceof List)){
            //list不为空，并且其中元素不是list处理excel
            List<?> objList = (List<?>) returnValue;
            if (!objList.isEmpty() && !(objList.get(0) instanceof List)) {
                excelWriteHandler.write(objList,responseExcel,response);
            }
        }

    }
}
