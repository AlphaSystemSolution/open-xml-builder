package com.alphasystem.openxml.builder.wml.table;

import org.apache.commons.lang3.ArrayUtils;

import java.math.BigDecimal;
import java.math.MathContext;
import java.util.ArrayList;
import java.util.List;

import static java.math.RoundingMode.HALF_UP;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * @author sali
 */
public final class ColumnAdapter {

    private static final BigDecimal TOTAL_GRID_COL_WIDTH = new BigDecimal(9576);
    private static final BigDecimal TOTAL_TABLE_WIDTH = new BigDecimal(5000);
    private static final BigDecimal PERCENT = new BigDecimal(100.0);
    private static final MathContext ROUNDING = new MathContext(2, HALF_UP);

    private final List<ColumnInfo> columns;
    private final BigDecimal totalTableWidth;

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
        BigDecimal _w = new BigDecimal(((totalTableWidthInPercent == null) || (totalTableWidthInPercent <= 0.0)) ?
                PERCENT.doubleValue() : totalTableWidthInPercent);
        BigDecimal totalGridWidth = TOTAL_GRID_COL_WIDTH.multiply(_w).divide(PERCENT);
        totalTableWidth = TOTAL_TABLE_WIDTH.multiply(_w).divide(PERCENT);

        final int length = isEmpty(columnWidths) ? 1 : columnWidths.length;
        columns = new ArrayList<>(length);

        BigDecimal[] widths = new BigDecimal[length];
        BigDecimal totalWidth = new BigDecimal(0.0);

        for (int i = 0; i < length; i++) {
            double columnWidth = columnWidths[i];
            final BigDecimal width = new BigDecimal(columnWidth, ROUNDING);
            widths[i] = width;
            totalWidth = totalWidth.add(width, ROUNDING);
        }

        for (int i = 0; i < length; i++) {
            BigDecimal columnWidthInPercent = widths[i].multiply(PERCENT).divide(totalWidth, ROUNDING);
            final double columnWidth = totalTableWidth.multiply(columnWidthInPercent).divide(PERCENT).doubleValue();
            final double gridWidth = totalGridWidth.multiply(columnWidthInPercent).divide(PERCENT).doubleValue();
            columns.add(new ColumnInfo(i, columnWidth, gridWidth));
        }
    }

    private static Double[] getColumnsWidthAsPercentages(int numOfColumns) {
        final BigDecimal width = PERCENT.divide(new BigDecimal(numOfColumns, ROUNDING));
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
}
