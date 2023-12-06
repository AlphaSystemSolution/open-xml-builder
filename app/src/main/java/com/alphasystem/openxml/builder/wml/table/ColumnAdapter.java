package com.alphasystem.openxml.builder.wml.table;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * @author sali
 */
public final class ColumnAdapter {

    private static final BigDecimal TOTAL_GRID_COL_WIDTH = TableAdapter.TOTAL_GRID_COL_WIDTH;
    private static final BigDecimal TOTAL_TABLE_WIDTH = TableAdapter.TOTAL_TABLE_WIDTH;
    private static final BigDecimal PERCENT = TableAdapter.PERCENT;
    private static final MathContext ROUNDING = TableAdapter.ROUNDING;

    private final BigDecimal totalTableWidth;
    private final List<ColumnInfo> columns;

    public ColumnAdapter(TableType tableType, int indentLevel, ColumnInput... columnWidthsInPercentage) {
        var columnInputs = isEmpty(columnWidthsInPercentage) ?
                new ColumnInput[]{new ColumnInput("col1", PERCENT.doubleValue())}
                : columnWidthsInPercentage;

        var numOfColumns = columnInputs.length;
        columns = new ArrayList<>(numOfColumns);

        var totalIndent = indentLevel >= 0 ? TableAdapter.DEFAULT_INDENT_VALUE + indentLevel * TableAdapter.DEFAULT_INDENT_VALUE : 0;
        // if there is indent then this would be subtracted from each column width
        final var indentPerColumn = BigDecimal.valueOf(totalIndent).divide(BigDecimal.valueOf(numOfColumns), ROUNDING);

        // populate actual columns info
        for (int index = 0; index < numOfColumns; index++) {
            columns.add(toColumnInfo(tableType, index, columnInputs[index], indentPerColumn.doubleValue()));
        }

        totalTableWidth = TableType.AUTO == tableType ? BigDecimal.ZERO : TOTAL_TABLE_WIDTH;
    }

    public List<ColumnInfo> getColumns() {
        return columns;
    }

    public BigDecimal getTotalTableWidth() {
        return totalTableWidth;
    }

    public ColumnInfo getColumn(int index) {
        return getColumns().get(index);
    }

    @Override
    public String toString() {
        return "ColumnAdapter{" +
                "totalTableWidth=" + totalTableWidth +
                ", columns=" + columns +
                '}';
    }

    private static ColumnInfo toColumnInfo(TableType tableType, int index, ColumnInput input, double indentPerColumn) {
        double width = input.getColumnWidth();
        var totalWidth = TableType.AUTO == tableType ? TOTAL_GRID_COL_WIDTH : TOTAL_TABLE_WIDTH;
        var columnWidth = totalWidth.multiply(BigDecimal.valueOf(width)).divide(PERCENT, ROUNDING)
                .setScale(0, RoundingMode.HALF_EVEN).doubleValue();
        columnWidth -= indentPerColumn;
        var gridWidth = TOTAL_GRID_COL_WIDTH.multiply(BigDecimal.valueOf(width)).divide(PERCENT, ROUNDING)
                .setScale(0, RoundingMode.HALF_EVEN).doubleValue();
        gridWidth = TableType.AUTO == tableType ? columnWidth : gridWidth;
        return new ColumnInfo(index, input.getColumnName(), columnWidth, gridWidth);
    }
}
