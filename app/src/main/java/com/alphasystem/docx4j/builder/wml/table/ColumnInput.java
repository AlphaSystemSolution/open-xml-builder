package com.alphasystem.docx4j.builder.wml.table;

public final class ColumnInput {

    private final String columnName;
    private final double columnWidth;

    public ColumnInput(String columnName, double columnWidth) {
        this.columnName = columnName;
        this.columnWidth = columnWidth;
    }

    public String getColumnName() {
        return columnName;
    }

    public double getColumnWidth() {
        return columnWidth;
    }

    @Override
    public String toString() {
        return "ColumnInput {" +
                "columnName='" + columnName + '\'' +
                ", columnWidth=" + columnWidth +
                " }";
    }
}
