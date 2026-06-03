package com.wxp.excel.listener;

import org.apache.fesod.sheet.read.listener.ReadListener;

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
