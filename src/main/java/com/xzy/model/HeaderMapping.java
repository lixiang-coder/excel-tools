package com.xzy.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 表头映射关系：多个源表头 -> 一个目标表头
 */
public class HeaderMapping {
    private final List<String> sourceHeaders;
    private final String targetHeader;
    private final FormatType formatType;

    public enum FormatType {
        COMMA_SEPARATED,    // 使用逗号分隔，不显示源表头
        PIPE_SEPARATED      // 使用竖线分隔，显示源表头
    }

    public HeaderMapping(List<String> sourceHeaders, String targetHeader) {
        this(sourceHeaders, targetHeader, FormatType.PIPE_SEPARATED);
    }

    public HeaderMapping(List<String> sourceHeaders, String targetHeader, FormatType formatType) {
        this.sourceHeaders = new ArrayList<>(sourceHeaders);
        this.targetHeader = targetHeader;
        this.formatType = formatType;
    }

    public List<String> getSourceHeaders() {
        return new ArrayList<>(sourceHeaders);
    }

    public String getTargetHeader() {
        return targetHeader;
    }

    public boolean isMerge() {
        return sourceHeaders.size() > 1;
    }

    public FormatType getFormatType() {
        return formatType;
    }
}
