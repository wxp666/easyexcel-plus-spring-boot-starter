package com.wxp.excel.strategy;

import com.alibaba.excel.write.handler.RowWriteHandler;
import com.alibaba.excel.write.handler.context.RowWriteHandlerContext;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Map;

@Data
@AllArgsConstructor
    public class ExcelMergeStrategy implements RowWriteHandler {

    //合并的起始行：key：开始，value；结束
    private Map<Integer, Integer> map;

    //要合并的列
    private int[] cols;

    @Override
    public void afterRowDispose(RowWriteHandlerContext context) {
        //如果是head或者当前行不是合并的起始行，跳过
        if (context.getHead() || !map.containsKey(context.getRowIndex())) {
            return;
        }
        Integer endrow = map.get(context.getRowIndex());
        if (!context.getRowIndex().equals(endrow)) {
            //编列合并的列，合并行
            for (int col : cols) {
                // CellRangeAddress(起始行,结束行,起始列,结束列)
                context.getWriteSheetHolder().getSheet().addMergedRegionUnsafe(new CellRangeAddress(context.getRowIndex(), endrow, col, col));
            }
        }
    }
}
