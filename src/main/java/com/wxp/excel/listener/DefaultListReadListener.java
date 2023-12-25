package com.wxp.excel.listener;

import com.alibaba.excel.context.AnalysisContext;

import java.util.ArrayList;
import java.util.List;

/**
 * @author wxp
 * @since 2023/12/22
 */
public class DefaultListReadListener extends ListReadListener<Object> {


    private final List<Object> dataList = new ArrayList<>();

    @Override
    public void invoke(Object data, AnalysisContext context) {
        dataList.add(data);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {

    }

    @Override
    public List<Object> getDataList() {
        return dataList;
    }

}
