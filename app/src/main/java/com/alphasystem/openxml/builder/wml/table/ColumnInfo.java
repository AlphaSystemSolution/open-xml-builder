package com.alphasystem.openxml.builder.wml.table;

/**
 * @author sali
 */
public final class ColumnInfo {

    private int columnNumber;
    private double columnWidth;
    private double gridWidth;

    public ColumnInfo() {
        this(0, 100.0, 100.0);
    }

    public ColumnInfo(int columnNumber, double columnWidth, double gridWidth) {
        setColumnNumber(columnNumber);
        setColumnWidth(columnWidth);
        setGridWidth(gridWidth);
    }

    public int getColumnNumber() {
        return columnNumber;
    }

    public void setColumnNumber(int columnNumber) {
        this.columnNumber = columnNumber;
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
}
