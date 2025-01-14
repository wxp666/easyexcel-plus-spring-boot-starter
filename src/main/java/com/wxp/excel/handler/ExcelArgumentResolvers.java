package com.wxp.excel.handler;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.annotation.ExcelProperty;
import com.wxp.excel.annotation.RequestExcel;
import com.wxp.excel.converters.LocalDateStringConverter;
import com.wxp.excel.converters.LocalDateTimeStringConverter;
import com.wxp.excel.exception.ExcelPlusException;
import com.wxp.excel.listener.ListReadListener;
import org.springframework.beans.BeanUtils;
import org.springframework.core.MethodParameter;
import org.springframework.core.ResolvableType;
import org.springframework.util.Assert;
import org.springframework.web.bind.support.WebDataBinderFactory;
import org.springframework.web.context.request.NativeWebRequest;
import org.springframework.web.method.support.HandlerMethodArgumentResolver;
import org.springframework.web.method.support.ModelAndViewContainer;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartRequest;

import javax.servlet.http.HttpServletRequest;
import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * @author wxp
 * @since 2023/12/22
 */
public class ExcelArgumentResolvers implements HandlerMethodArgumentResolver {

    @Override
    public boolean supportsParameter(MethodParameter parameter) {
        // 判断方法是否标注RequestExcel注解
        return parameter.hasParameterAnnotation(RequestExcel.class);
    }

    @Override
    public Object resolveArgument(MethodParameter parameter, ModelAndViewContainer mavContainer, NativeWebRequest webRequest, WebDataBinderFactory binderFactory) throws Exception {
        Class<?> parameterType = parameter.getParameterType();
        if (!parameterType.isAssignableFrom(List.class)) {
            throw new ExcelPlusException(
                    "excel 上传解析错误, 参数不是一个list" + parameterType);
        }

        // 处理自定义 readListener
        RequestExcel requestExcel = parameter.getParameterAnnotation(RequestExcel.class);
        Assert.notNull(requestExcel, "Excel导入RequestExcel为空");
        Class<? extends ListReadListener<?>> readListenerClass = requestExcel.readListener();
        ListReadListener<?> readListener = BeanUtils.instantiateClass(readListenerClass);

        // 获取请求文件流
        HttpServletRequest request = webRequest.getNativeRequest(HttpServletRequest.class);
        Assert.notNull(request, "Excel导入HttpServletRequest为空");
        InputStream inputStream;
        if (request instanceof MultipartRequest) {
            MultipartFile file = ((MultipartRequest) request).getFile(requestExcel.fileName());
            Assert.notNull(file, "excel导入文件不能为空");
            inputStream = file.getInputStream();
        } else {
            inputStream = request.getInputStream();
        }

        // 获取目标类型
        Class<?> excelModelClass = ResolvableType.forMethodParameter(parameter).getGeneric(0).resolve();
        Assert.notNull(excelModelClass, "excel导入目标类型为空");
        // head行数
        int headNumber = requestExcel.headNumber();
        if (requestExcel.headNumber() == -1) {
            Field excelPropertyField = Arrays.stream(excelModelClass.getDeclaredFields()).filter(field -> field.getAnnotation(ExcelProperty.class) != null)
                    .findFirst().orElseThrow(() -> new ExcelPlusException("excel上传模板类未找到ExcelProperty注解字段"));
            ExcelProperty excelProperty = excelPropertyField.getAnnotation(ExcelProperty.class);
            String[] value = excelProperty.value();
            headNumber = value.length == 0 ? 1 : value.length;
        }

        EasyExcel.read(inputStream, excelModelClass, readListener)
                .registerConverter(new LocalDateStringConverter())
                .registerConverter(new LocalDateTimeStringConverter())
                .sheet(0)
                .headRowNumber(headNumber)
                .doRead();
        if (requestExcel.valid()) {
            List dataList = readListener.getDataList();
            Validator validator = Validation.buildDefaultValidatorFactory().getValidator();
            List<String> errorMessages = new ArrayList<>();
            for (int i = 0; i < dataList.size(); i++) {
                Set<ConstraintViolation<Object>> validate = validator.validate(dataList.get(i));
                if (!validate.isEmpty()) {
                    errorMessages.add("第" + (i + 1 + headNumber) + "行数据校验失败:" + validate.stream().map(ConstraintViolation::getMessage).collect(Collectors.joining(",")));
                }
            }
            if (!errorMessages.isEmpty()) {
                throw new ExcelPlusException(String.join(";", errorMessages));
            }
        }
        return readListener.getDataList();
    }
}
