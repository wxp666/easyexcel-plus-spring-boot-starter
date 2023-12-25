package com.wxp.excel.listener;

import com.alibaba.excel.read.listener.ReadListener;
import com.wxp.excel.exception.ErrorMessage;

import java.util.List;

/**
 * @author wxp
 * @since 2023/12/22
 */
public abstract class ListReadListener<T> implements ReadListener<T> {

    /**
     * 获取数据列表
     *
     * @return data
     */
    public abstract List<T> getDataList();

}
