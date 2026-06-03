package com.wxp.excel.strategy;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.fesod.common.util.MapUtils;
import org.apache.fesod.sheet.enums.CellDataTypeEnum;
import org.apache.fesod.sheet.metadata.Head;
import org.apache.fesod.sheet.metadata.data.WriteCellData;
import org.apache.fesod.sheet.write.metadata.holder.WriteSheetHolder;
import org.apache.fesod.sheet.write.style.column.AbstractColumnWidthStyleStrategy;
import org.apache.poi.ss.usermodel.Cell;

import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author wxp
 * @date 2023/3/13
 * @apiNote
 */
public class AutoColumnWidthStrategy extends AbstractColumnWidthStyleStrategy {
    private static final int MAX_COLUMN_WIDTH = 255;
    private final Map<Integer, Map<Integer, Integer>> cache = MapUtils.newHashMapWithExpectedSize(8);

    public AutoColumnWidthStrategy() {}

    @Override
    protected void setColumnWidth(WriteSheetHolder writeSheetHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        boolean needSetWidth = isHead || !CollectionUtils.isEmpty(cellDataList);
        if (needSetWidth) {
            Map<Integer, Integer> maxColumnWidthMap = this.cache.computeIfAbsent(writeSheetHolder.getSheetNo(), key -> HashMap. newHashMap(16));
            Integer columnWidth = this.dataLength(cellDataList, cell, isHead);
            if (columnWidth > 0) {
                if (columnWidth > MAX_COLUMN_WIDTH) {
                    columnWidth = MAX_COLUMN_WIDTH;
                }
                Integer maxColumnWidth = maxColumnWidthMap.get(cell.getColumnIndex());
                if (maxColumnWidth == null || columnWidth > maxColumnWidth) {
                    maxColumnWidthMap.put(cell.getColumnIndex(), columnWidth);
                    writeSheetHolder.getSheet().setColumnWidth(cell.getColumnIndex(), columnWidth * 256);
                }
            }
        }
    }

    private Integer dataLength(List<WriteCellData<?>> cellDataList, Cell cell, Boolean isHead) {
        if (isHead) {
            return cell.getStringCellValue().getBytes(StandardCharsets.UTF_8).length;
        }
        WriteCellData<?> cellData = cellDataList.get(0);
        CellDataTypeEnum type = cellData.getType();
        if (type == null) {
            return -1;
        } else {
            switch (type) {
                case STRING:
                    return cellData.getStringValue().getBytes(StandardCharsets.UTF_8).length + 2;
                case BOOLEAN:
                    return cellData.getBooleanValue().toString().getBytes(StandardCharsets.UTF_8).length;
                case NUMBER:
                    return cellData.getNumberValue().toString().getBytes(StandardCharsets.UTF_8).length;
                case DATE:
                    return cellData.getDateValue().toString().getBytes(StandardCharsets.UTF_8).length + 2;
                default:
                    return -1;
            }
        }

    }

}
