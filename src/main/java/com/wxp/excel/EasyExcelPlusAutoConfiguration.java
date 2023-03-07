package com.wxp.excel;

import com.wxp.excel.handler.ExcelReturnValueHandler;
import com.wxp.excel.handler.ExcelWriteHandler;
import lombok.RequiredArgsConstructor;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.annotation.Bean;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.mvc.method.annotation.RequestMappingHandlerAdapter;

import javax.annotation.PostConstruct;
import java.util.ArrayList;
import java.util.List;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote 自动配置类
 */
@AutoConfiguration
@RequiredArgsConstructor
public class EasyExcelPlusAutoConfiguration {

    private final RequestMappingHandlerAdapter requestMappingHandlerAdapter;

    @Bean
    @ConditionalOnMissingBean
    public ExcelWriteHandler excelWriteHandler(){
        return new ExcelWriteHandler();
    }

    @Bean
    @ConditionalOnMissingBean
    public ExcelReturnValueHandler excelReturnValueHandler(ExcelWriteHandler excelWriteHandler){
        return new ExcelReturnValueHandler(excelWriteHandler);
    }

    /**
     * 追加处理器到springmvc
     */
    @PostConstruct
    public void setReturnValueHandlers() {
        List<HandlerMethodReturnValueHandler> returnValueHandlers = requestMappingHandlerAdapter
                .getReturnValueHandlers();

        List<HandlerMethodReturnValueHandler> newHandlers = new ArrayList<>();
        newHandlers.add(excelReturnValueHandler(excelWriteHandler()));
        assert returnValueHandlers != null;
        newHandlers.addAll(returnValueHandlers);
        requestMappingHandlerAdapter.setReturnValueHandlers(newHandlers);
    }

}
