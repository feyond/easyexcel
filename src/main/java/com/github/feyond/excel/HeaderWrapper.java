package com.github.feyond.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;

/**
 * @author chenfy
 * @create 2017-10-23 9:59
 **/
@Data
@NoArgsConstructor
@AllArgsConstructor
public class HeaderWrapper {
    private String header;
    private String comment;
    public static final String COMMENT_SEPARATOR = "**";

    public HeaderWrapper(String header) {
        this(getRealHeader(header), getRealComment(header));
    }

    private static String getRealHeader(String header) {
        return StringUtils.split(header, COMMENT_SEPARATOR)[0];
    }

    private static String getRealComment(String header) {
        if (StringUtils.contains(header, COMMENT_SEPARATOR)) {
            return StringUtils.split(header, COMMENT_SEPARATOR)[1];
        }
        return StringUtils.EMPTY;
    }
}
