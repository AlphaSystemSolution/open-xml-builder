package com.alphasystem.openxml.builder.wml.table;

/**
 * @author sali
 */
public final class ColumnInfo {

    private int columnNumber;
    private String columnName;
    private double columnWidth;
    private double gridWidth;

    public ColumnInfo() {
        this(0,  100.0, 100.0);
    }

    public ColumnInfo(int columnNumber, double columnWidth, double gridWidth) {
        this(columnNumber, "", columnWidth, gridWidth);
    }

    public ColumnInfo(int columnNumber, String columnName, double columnWidth, double gridWidth) {
        setColumnNumber(columnNumber);
        setColumnName(columnName);
        setColumnWidth(columnWidth);
        setGridWidth(gridWidth);
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public void setColumnNumber(int columnNumber) {
        this.columnNumber = columnNumber;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }


    public double getColumnWidth() {
        return columnWidth;
    }

    public void setColumnWidth(double columnWidth) {
        this.columnWidth = columnWidth;
    }

    public double getGridWidth() {
        return gridWidth;
    }

    public void setGridWidth(double gridWidth) {
        this.gridWidth = gridWidth;
    }

    @Override
    public String toString() {
        return "ColumnInfo{" +
                "columnNumber=" + columnNumber +
                ", columnName='" + columnName + '\'' +
                ", columnWidth=" + columnWidth +
                ", gridWidth=" + gridWidth +
                '}';
    }
}
