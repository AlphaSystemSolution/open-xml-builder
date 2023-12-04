package com.alphasystem.openxml.builder.wml.table;

import org.apache.commons.lang3.ArrayUtils;

import java.math.BigDecimal;
import java.math.MathContext;
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

    public ColumnAdapter(BigDecimal totalTableWidth, List<ColumnInfo> columns) {
        this.totalTableWidth = totalTableWidth;
        this.columns = columns;
    }

    /**
     * Constructor to create AUTO table.
     *
     * @param numOfColumns number of columns
     * @param indentLevel  indent level
     */
    public ColumnAdapter(int numOfColumns, int indentLevel) {
        totalTableWidth = BigDecimal.ZERO;
        var totalIndent = indentLevel >= 0 ? TableAdapter.DEFAULT_INDENT_VALUE + indentLevel * TableAdapter.DEFAULT_INDENT_VALUE : 0;
        var indentPerColumn = totalIndent / numOfColumns;
        var individualColumnWidth = TOTAL_GRID_COL_WIDTH.divide(BigDecimal.valueOf(numOfColumns), ROUNDING);
        var columnWidths = new ArrayList<Integer>();
        for (int i = 0; i < numOfColumns; i++) {
            columnWidths.add(individualColumnWidth.intValue());
        }
        var sum = columnWidths.stream().reduce(0, Integer::sum);
        var diff = TOTAL_GRID_COL_WIDTH.subtract(BigDecimal.valueOf(sum));
        var index = 0;
        while (diff.intValue() > 0) {
            columnWidths.set(index, columnWidths.get(index) + 1);
            index += 1;
            diff = diff.subtract(BigDecimal.ONE);
        }

        columns = new ArrayList<>(numOfColumns);
        for (int i = 0; i < columnWidths.size(); i++) {
            var columnWidth = columnWidths.get(i) - indentPerColumn;
            columns.add(new ColumnInfo(i, columnWidth, columnWidth));
        }
    }

    public ColumnAdapter(int numOfColumns) {
        this(PERCENT.doubleValue(), numOfColumns);
    }

    public ColumnAdapter(Double tableWidthInPercent, int numOfColumns) {
        this(tableWidthInPercent, getColumnsWidthAsPercentages(numOfColumns));
    }

    public ColumnAdapter(Double... columnWidthPercentages) {
        this(PERCENT.doubleValue(), columnWidthPercentages);
    }

    public ColumnAdapter(Double totalTableWidthInPercent, Double... columnWidths) {
        BigDecimal _w = BigDecimal.valueOf(((totalTableWidthInPercent == null) || (totalTableWidthInPercent <= 0.0)) ?
                PERCENT.doubleValue() : totalTableWidthInPercent);
        BigDecimal totalGridWidth = TOTAL_GRID_COL_WIDTH.multiply(_w).divide(PERCENT, ROUNDING);
        totalTableWidth = TOTAL_TABLE_WIDTH.multiply(_w).divide(PERCENT, ROUNDING);

        final int length = isEmpty(columnWidths) ? 1 : columnWidths.length;
        columns = new ArrayList<>(length);

        BigDecimal[] widths = new BigDecimal[length];
        BigDecimal totalWidth = BigDecimal.valueOf(0.0);

        for (int i = 0; i < length; i++) {
            double columnWidth = columnWidths[i];
            final BigDecimal width = new BigDecimal(columnWidth, ROUNDING);
            widths[i] = width;
            totalWidth = totalWidth.add(width, ROUNDING);
        }

        for (int i = 0; i < length; i++) {
            BigDecimal columnWidthInPercent = widths[i].multiply(PERCENT).divide(totalWidth, ROUNDING);
            final double columnWidth = totalTableWidth.multiply(columnWidthInPercent).divide(PERCENT, ROUNDING).doubleValue();
            final double gridWidth = totalGridWidth.multiply(columnWidthInPercent).divide(PERCENT, ROUNDING).doubleValue();
            columns.add(new ColumnInfo(i, columnWidth, gridWidth));
        }
    }

    private static Double[] getColumnsWidthAsPercentages(int numOfColumns) {
        final BigDecimal width = PERCENT.divide(BigDecimal.valueOf(numOfColumns), ROUNDING);
        Double[] columnWidthPercentages = new Double[numOfColumns];
        return ArrayUtils.add(columnWidthPercentages, width.doubleValue());
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
}
