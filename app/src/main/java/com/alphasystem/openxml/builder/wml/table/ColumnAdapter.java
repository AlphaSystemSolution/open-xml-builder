package com.alphasystem.openxml.builder.wml.table;

import org.apache.commons.lang3.ArrayUtils;

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

    private static final BigDecimal TOTAL_GRID_COL_WIDTH = BigDecimal.valueOf(9576);
    private static final BigDecimal TOTAL_TABLE_WIDTH = BigDecimal.valueOf(5000);
    private static final BigDecimal PERCENT = BigDecimal.valueOf(100.0);
    private static final MathContext ROUNDING = new MathContext(4, RoundingMode.CEILING);

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
}
