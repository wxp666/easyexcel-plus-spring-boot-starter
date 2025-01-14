package com.wxp.excel.annotation;

import com.wxp.excel.listener.DefaultListReadListener;
import com.wxp.excel.listener.ListReadListener;

import java.lang.annotation.*;

/**
 * @author wxp
 * @since 2023/12/22
 */
@Documented
@Target({ ElementType.PARAMETER })
@Retention(RetentionPolicy.RUNTIME)
public @interface RequestExcel {

    /**
     * 接口上传文件参数名称
     *
     * @return 参数名称
     */
    String fileName();

    /**
     * 读取监听器
     *
     * @return 监听器
     */
    Class<? extends ListReadListener<?>> readListener() default DefaultListReadListener.class;

    /**
     * 表头行数
     * 不指定默认-1
     * 取一个ExcelProperty注解 value有值时 取value的长度
     * 如value没有数值则 head 取值为 1
     *
     * @return 表头行数
     */
    int headNumber() default -1;

    /**
     * 是否校验数据
     * 校验validation相关注解
     * @return 是否校验
     */
    boolean valid() default false;


}
