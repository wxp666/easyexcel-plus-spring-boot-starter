package com.wxp.excel.exception;

import com.alibaba.excel.exception.ExcelRuntimeException;

/**
 * @author wxp
 * @date 2023/3/13
 * @apiNote
 */
public class ExcelPlusException extends ExcelRuntimeException {

    public ExcelPlusException() {
    }

    public ExcelPlusException(String message) {
        super(message);
    }

    public ExcelPlusException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelPlusException(Throwable cause) {
        super(cause);
    }
}
