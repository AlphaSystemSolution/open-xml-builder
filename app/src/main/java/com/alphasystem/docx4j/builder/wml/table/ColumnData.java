package com.alphasystem.docx4j.builder.wml.table;

import org.docx4j.wml.TcPr;

import java.util.Objects;

public final class ColumnData {

    private final Integer columnIndex;
    private Integer gridSpanValue = null;
    private VerticalMergeType verticalMergeType = null;
    private TcPr columnProperties = null;
    private Object[] content = new Object[0];

    public ColumnData(Integer columnIndex) {
        Objects.requireNonNull(columnIndex, "columnIndex cannot be null");
        this.columnIndex = columnIndex;
    }

    public ColumnData withGridSpanValue(Integer gridSpanValue) {
        this.gridSpanValue = gridSpanValue;
        return this;
    }

    public ColumnData withVerticalMergeType(VerticalMergeType verticalMergeType) {
        this.verticalMergeType = verticalMergeType;
        return this;
    }

    public ColumnData withColumnProperties(TcPr columnProperties) {
        this.columnProperties = columnProperties;
        return this;
    }

    public ColumnData withContent(Object... content) {
        this.content = content;
        return this;
    }

    public Integer getColumnIndex() {
        return columnIndex;
    }

    public Integer getGridSpanValue() {
        return gridSpanValue;
    }

    public VerticalMergeType getVerticalMergeType() {
        return verticalMergeType;
    }

    public TcPr getColumnProperties() {
        return columnProperties;
    }

    public Object[] getContent() {
        return content;
    }
}
