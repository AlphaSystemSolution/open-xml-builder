package com.alphasystem.openxml.builder.wml.table;

import com.alphasystem.openxml.builder.wml.*;
import org.docx4j.wml.*;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.sharedtypes.STOnOff.ONE;
import static org.docx4j.wml.TblWidth.TYPE_DXA;

/**
 * @author sali
 */
public final class TableAdapter {

    public static final BigDecimal TOTAL_GRID_COL_WIDTH = BigDecimal.valueOf(9576);
    public static final BigDecimal TOTAL_TABLE_WIDTH = BigDecimal.valueOf(5000);
    public static final int DEFAULT_INDENT_VALUE = 720;
    public static final BigDecimal PERCENT = BigDecimal.valueOf(100.0);
    public static final String DEFAULT_TABLE_STYLE = "TableGrid";
    public static final MathContext ROUNDING = new MathContext(4, RoundingMode.CEILING);

    private final TableType tableType;
    private ColumnAdapter columnAdapter;
    private TblBuilder tblBuilder;
    private TrBuilder trBuilder;

    public TableAdapter() {
        this(TableType.PCT);
    }

    public TableAdapter(final TableType tableType) {
        this.tableType = tableType;
    }

    public TableAdapter startTable(int numOfColumns) {
        return startTable(getColumnWidths(numOfColumns).toArray(new Double[0]));
    }

    public TableAdapter startTable(Double... columnWidths) {
        return startTable((TblPr) null, columnWidths);
    }

    public TableAdapter startTable(TblPr tableProperties, Double... columnWidths) {
        return startTable(null, null, -1, tableProperties, columnWidths);
    }

    public TableAdapter startTable(String tableStyle, Double... columnWidths) {
        return startTable(tableStyle, null, columnWidths);
    }

    public TableAdapter startTable(String tableStyle, TblPr tableProperties, Double... columnWidths) {
        return startTable(null, tableStyle, -1, tableProperties, columnWidths);
    }

    public TableAdapter startTable(ColumnAdapter columnAdapter,
                                   String tableStyle,
                                   int indentLevel,
                                   TblPr tableProperties) {
        tblBuilder = getTblBuilder();
        this.columnAdapter = columnAdapter;

        TblGridBuilder tblGridBuilder = getTblGridBuilder();
        columnAdapter.getColumns().forEach(columnInfo -> {
            final TblGridCol tblGridCol = getTblGridColBuilder().withW((long) columnInfo.getGridWidth()).getObject();
            tblGridBuilder.addGridCol(tblGridCol);
        });

        TblWidth tblWidth = getTblWidthBuilder().withType(tableType.getTableType())
                .withW(columnAdapter.getTotalTableWidth().longValue()).getObject();

        var cTTblLook = getCTTblLookBuilder().withFirstRow(ONE).withLastRow(ONE).withFirstColumn(ONE)
                .withLastColumn(ONE).withNoVBand(ONE).withNoHBand(ONE).getObject();

        tableStyle = isBlank(tableStyle) ? DEFAULT_TABLE_STYLE : tableStyle;
        TblWidth tblIndent = null;
        if (indentLevel >= 0) {
            long indentValue = DEFAULT_INDENT_VALUE + ((long) indentLevel * DEFAULT_INDENT_VALUE);
            tblIndent = getTblWidthBuilder().withType(TYPE_DXA).withW(indentValue).getObject();
        }
        TblPr tblPr = getTblPrBuilder().withTblStyle(tableStyle).withTblW(tblWidth).withTblInd(tblIndent)
                .withTblLook(cTTblLook).getObject();
        TblPrBuilder tblPrBuilder = new TblPrBuilder(tblPr, tableProperties);

        tblBuilder.withTblGrid(tblGridBuilder.getObject()).withTblPr(tblPrBuilder.getObject());
        return this;
    }

    private TableAdapter startTable(String tableStyle,
                                    int numOfColumns,
                                    int indentLevel,
                                    Double totalTableWidthInPercent,
                                    TblPr tableProperties,
                                    Double... columnWidths) {
        ColumnAdapter columnAdapter;
        if (tableType == TableType.PCT) {
            columnAdapter = new ColumnAdapter(totalTableWidthInPercent, columnWidths);
        } else {
            columnAdapter = new ColumnAdapter(numOfColumns, indentLevel);
        }
        return startTable(columnAdapter, tableStyle,indentLevel, tableProperties);
    }

    public TableAdapter startAutoTable(String tableStyle, int numOfColumns, int indentLevel, TblPr tableProperties) {
        return startTable(tableStyle, numOfColumns, indentLevel, 0.0, tableProperties);
    }


    public TableAdapter startAutoTable(String tableStyle, int numOfColumns, int indentLevel) {
        return startTable(tableStyle, numOfColumns, indentLevel, 0.0, null);
    }

    public TableAdapter startAutoTable(int numOfColumns, int indentLevel, TblPr tableProperties) {
        return startTable(null, numOfColumns, indentLevel, 0.0, tableProperties);
    }

    public TableAdapter startAutoTable(int numOfColumns, int indentLevel) {
        return startTable(null, numOfColumns, indentLevel, 0.0, null);
    }

    public TableAdapter startTable(Double totalTableWidthInPercent, String tableStyle, int indentLevel, TblPr tableProperties,
                                   Double... columnWidths) {
        return startTable(tableStyle, 0, indentLevel, totalTableWidthInPercent, tableProperties, columnWidths);
    }

    public TableAdapter startRow() {
        return startRow(getTrBuilder().withRsidR(nextId()).withRsidTr(nextId()));
    }

    public TableAdapter startRow(TrBuilder trBuilder) {
        this.trBuilder = trBuilder;
        return this;
    }

    public TableAdapter endRow() {
        tblBuilder.addContent(trBuilder.getObject());
        trBuilder = null;
        return this;
    }

    public TableAdapter addColumn(Integer columnIndex, Object... content) {
        return addColumn(columnIndex, null, null, content);
    }

    public TableAdapter addColumn(Integer columnIndex, TcPr columnProperties, Object... content) {
        return addColumn(columnIndex, null, columnProperties, content);
    }

    public TableAdapter addColumn(Integer columnIndex, Integer gridSpanValue, Object... content) {
        return addColumn(columnIndex, gridSpanValue, null, content);
    }

    public TableAdapter addColumn(Integer columnIndex, Integer gridSpanValue, TcPr columnProperties, Object... content) {
        return addColumn(columnIndex, gridSpanValue, null, columnProperties, content);
    }

    public TableAdapter addColumn(Integer columnIndex, Integer gridSpanValue, VerticalMergeType verticalMergeType,
                                  TcPr columnProperties, Object... content) {
        final Tc tc = getTcBuilder().withTcPr(getColumnProperties(columnAdapter, tableType, columnIndex, gridSpanValue,
                verticalMergeType, columnProperties)).addContent(content).getObject();
        trBuilder.addContent(tc);
        return this;
    }

    public Tbl getTable() {
        return tblBuilder.getObject();
    }


    public long getTotalTableWidth() {
        return columnAdapter == null ? 0L : columnAdapter.getTotalTableWidth().longValue();
    }

    private static TcPr getColumnProperties(ColumnAdapter columnAdapter,
                                            TableType tableType,
                                            Integer columnIndex,
                                            Integer gridSpanValue,
                                            VerticalMergeType verticalMergeType,
                                            TcPr columnProperties)
            throws ArrayIndexOutOfBoundsException {
        List<ColumnInfo> columns = columnAdapter.getColumns();
        checkColumnIndex(columns, columnIndex);
        final ColumnInfo columnInfo = columnAdapter.getColumn(columnIndex);
        BigDecimal columnWidth = BigDecimal.valueOf(columnInfo.getColumnWidth());
        long gs = 1;
        if (gridSpanValue != null && gridSpanValue > 1) {
            // sanity check, make sure we are not going out of bound
            checkColumnIndex(columns, columnIndex + gridSpanValue - 1);
            // iterate through width and get the total width for the grid span
            for (int i = columnIndex + 1; i < columnIndex + gridSpanValue; i++) {
                final ColumnInfo columnInfo1 = columns.get(i);
                columnWidth = columnWidth.add(BigDecimal.valueOf(columnInfo1.getColumnWidth()));
            }
            gs = gridSpanValue;
        }

        TcPrBuilder tcPrBuilder = getTcPrBuilder();
        TcPrInner.VMerge vMerge = null;
        if (verticalMergeType != null) {
            vMerge = tcPrBuilder.getVMergeBuilder().withVal(verticalMergeType.getValue()).getObject();
        }

        TblWidth tblWidth = getTblWidthBuilder().withType(tableType.getColumnType()).withW(columnWidth.longValue()).getObject();
        tcPrBuilder.withGridSpan(gs).withTcW(tblWidth).withVMerge(vMerge);

        return new TcPrBuilder(tcPrBuilder.getObject(), columnProperties).getObject();
    }

    // private methods

    /**
     * Checks whether <code>columnIndex</code> is within range.
     *
     * @param columnInfos column data
     * @param columnIndex index of column
     * @throws ArrayIndexOutOfBoundsException if <code>columnIndex</code> is out of bound.
     */
    private static void checkColumnIndex(List<ColumnInfo> columnInfos, int columnIndex)
            throws ArrayIndexOutOfBoundsException {
        int numOfColumns = columnInfos.size();
        if (columnIndex < 0 || columnIndex >= numOfColumns) {
            throw new ArrayIndexOutOfBoundsException(
                    format("Invalid columnIndex {%s}, expected values are between %s and %s",
                            columnIndex, 0, (numOfColumns - 1)));
        }
    }

    private static List<Double> getColumnWidths(int numOfColumns) {
        var singleColumnWidth = PERCENT.divide(BigDecimal.valueOf(numOfColumns), ROUNDING);
        var columnWidths = new ArrayList<Double>(numOfColumns);
        for (int i = 0; i < numOfColumns; i++) {
            columnWidths.add(singleColumnWidth.doubleValue());
        }

        var sum = columnWidths.stream().reduce(0.0, Double::sum);
        var diff = PERCENT.subtract(BigDecimal.valueOf(sum));
        var index = 0;
        while (diff.intValue() > 0) {
            columnWidths.set(index, columnWidths.get(index) + 1);
            index += 1;
            diff = diff.subtract(BigDecimal.ONE);
        }

        return columnWidths;
    }

    public enum VerticalMergeType {

        RESTART("restart"), CONTINUE(null);

        private final String value;

        VerticalMergeType(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }
    }
}

