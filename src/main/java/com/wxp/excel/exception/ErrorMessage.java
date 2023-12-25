package com.wxp.excel.exception;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Set;

/**
 * @author wxp
 * @since 2023/12/22
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class ErrorMessage {

    private Integer lineNum;

    private Set<String> messages;

}
