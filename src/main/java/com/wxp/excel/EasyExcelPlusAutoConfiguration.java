package com.wxp.excel;

import com.wxp.excel.handler.ExcelArgumentResolvers;
import com.wxp.excel.handler.ExcelReturnValueHandler;
import com.wxp.excel.handler.ExcelWriteHandler;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.context.annotation.Bean;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.HandlerMethodReturnValueHandler;
import org.springframework.web.servlet.config.annotation.WebMvcConfigurer;

import java.util.List;

/**
 * @author wxp
 * @date 2023/3/3
 * @apiNote 自动配置类
 */
@AutoConfiguration
public class EasyExcelPlusAutoConfiguration implements WebMvcConfigurer {

    @Bean
    @ConditionalOnMissingBean
    public ExcelWriteHandler excelWriteHandler() {
        return new ExcelWriteHandler();
    }

    @Bean
    @ConditionalOnMissingBean
    public ExcelReturnValueHandler excelReturnValueHandler(ExcelWriteHandler excelWriteHandler) {
        return new ExcelReturnValueHandler(excelWriteHandler);
    }

    @Bean
    @ConditionalOnMissingBean
    public ExcelArgumentResolvers excelArgumentResolvers() {
        return new ExcelArgumentResolvers();
    }

    @Override
    public void addReturnValueHandlers(List<HandlerMethodReturnValueHandler> handlers) {
        handlers.add(0, excelReturnValueHandler(excelWriteHandler()));
    }

    @Override
    public void addArgumentResolvers(List<HandlerMethodArgumentResolver> resolvers) {
        resolvers.add(excelArgumentResolvers());
    }
}
