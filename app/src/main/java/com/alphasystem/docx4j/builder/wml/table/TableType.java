package com.alphasystem.docx4j.builder.wml.table;

public enum TableType {
    AUTO("auto", "dxa"), PCT("pct", "pct");

    private final String tableType;
    private final String columnType;

    TableType(String value, String columType) {
        this.tableType = value;
        this.columnType = columType;
    }

    public String getTableType() {
        return tableType;
    }

    public String getColumnType() {
        return columnType;
    }
}
