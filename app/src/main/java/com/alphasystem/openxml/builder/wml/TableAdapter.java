/**
 *
 */
package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.*;
import org.docx4j.wml.TcPrInner.GridSpan;

import java.math.BigDecimal;
import java.math.MathContext;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getGridSpan;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.math.RoundingMode.CEILING;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.sharedtypes.STOnOff.ONE;

/**
 * Table Adapter.
 * <p>
 * <p>
 * <pre>
 *       <code>TableAdapter tableAdapter = new TableAdapter(3)</code>
 *       <code>tableAdapter.setColumnWidth(0, 20.0).setColumnWidth(1, 30.0).setColumnWidth(3, 50.0);</code>
 *       <code>tableAdapter.startTable().startRow();</code>
 *       <code>// add column(s)</code>
 *       <code>tableAdapter.endRow();</code>
 *       <code>// make use of Tbl</code>
 *       <code>Tbl tbl = tableAdapter.getTable();</code>
 * </pre>
 * <p>
 * </p>
 *
 * @author sali
 */
@Deprecated
public class TableAdapter {

    private static final String TYPE_PCT = "pct";

    private static final BigDecimal TOTAL_GRID_COL_WIDTH = new BigDecimal(9576);

    private static final BigDecimal TOTAL_TABLE_WIDTH = new BigDecimal(5000);

    private static final BigDecimal PERCENT = new BigDecimal(100.0);

    private static final MathContext ROUNDING = new MathContext(3, CEILING);
    private final int numOfColumns;
    private final BigDecimal bigNumOfColumn;
    private final Integer[] gridWidths;
    private final Integer[] columnsWidths;
    private final double defaultPercent;
    private TblBuilder tblBuilder;
    private TrBuilder trBuilder;

    /**
     * @param columnWidthPercentages array of width of columns, <strong>this should be percentage of columns width</strong>
     * @throws IllegalArgumentException if argument <code>columnWidthPercentages</code> is null or array of length 0
     */
    public TableAdapter(Double... columnWidthPercentages)
            throws IllegalArgumentException {
        if (isEmpty(columnWidthPercentages)) {
            throw new IllegalArgumentException("Columns width are not be initialized");
        }
        tblBuilder = getTblBuilder();
        numOfColumns = columnWidthPercentages.length;
        this.bigNumOfColumn = new BigDecimal(numOfColumns);
        this.gridWidths = new Integer[numOfColumns];
        this.columnsWidths = new Integer[numOfColumns];
        this.defaultPercent = getDefaultPercent();

        for (int index = 0; index < columnWidthPercentages.length; index++) {
            Double percent = columnWidthPercentages[index];
            setColumnWidthInternal(index, percent);
        }
    }

    /**
     * Create <code>TableAdapter</code> with specified number of columns.
     *
     * @param numOfColumns total number of columns
     */
    public TableAdapter(int numOfColumns) {
        this(new Double[numOfColumns]);
    }

    /**
     * Add a column into this table at specified index.
     *
     * @param columnIndex   column index
     * @param gridSpanValue value for this column to span
     * @param tcPr          table column properties
     * @param content       content of this column
     * @return reference to this
     * @throws ArrayIndexOutOfBoundsException if <code>columnIndex</code> is out of bound
     */
    public TableAdapter addColumn(Integer columnIndex, Integer gridSpanValue, TcPr tcPr, Object... content) throws
            ArrayIndexOutOfBoundsException {
        checkColumnIndex(columnIndex);
        Integer cw = columnsWidths[columnIndex];
        GridSpan gridSpan = getGridSpan(1);
        if (gridSpanValue != null && gridSpanValue > 1) {
            // sanity check, make sure we are not going out of bound
            checkColumnIndex(columnIndex + gridSpanValue - 1);
            // iterate through width and get the total width for the grid span
            for (int i = columnIndex + 1; i < gridSpanValue; i++) {
                cw += columnsWidths[i];
            }
            gridSpan = getGridSpan(gridSpanValue);
        }
        TblWidth tblWidth = getTblWidthBuilder().withType(TYPE_PCT).withW(cw.toString()).getObject();
        TcPrBuilder tcPrBuilder = getTcPrBuilder().withGridSpan(gridSpan).withTcW(tblWidth);
        tcPrBuilder = new TcPrBuilder(tcPrBuilder.getObject(), tcPr);
        Tc tc = getTcBuilder().withTcPr(tcPrBuilder.getObject()).addContent(content).getObject();
        trBuilder.addContent(tc);
        return this;
    }

    /**
     * Add a column into this table at specified index.
     *
     * @param columnIndex column index
     * @param content     content of this column
     * @return reference to this
     * @throws ArrayIndexOutOfBoundsException if <code>columnIndex</code> is out of bound
     */
    public TableAdapter addColumn(Integer columnIndex, Object... content) throws ArrayIndexOutOfBoundsException {
        return addColumn(columnIndex, null, null, content);
    }

    /**
     * Checks whether <code>columnIndex</code> is within range.
     *
     * @param columnIndex index of column
     * @throws ArrayIndexOutOfBoundsException if <code>columnIndex</code> is out of bound.
     */
    protected void checkColumnIndex(int columnIndex)
            throws ArrayIndexOutOfBoundsException {
        if (columnIndex < 0 || columnIndex >= numOfColumns) {
            throw new ArrayIndexOutOfBoundsException(
                    format("Invalid columnIndex {%s}, expected values are between %s and %s",
                            columnIndex, 0, (numOfColumns - 1)));
        }
    }

    /**
     * Ends current row.
     *
     * @return reference to this
     */
    public TableAdapter endRow() {
        tblBuilder.addContent(trBuilder.getObject());
        trBuilder = null;
        return this;
    }

    private double getDefaultPercent() {
        BigDecimal width = TOTAL_TABLE_WIDTH.divide(bigNumOfColumn, ROUNDING);
        return width.divide(TOTAL_TABLE_WIDTH).multiply(PERCENT, ROUNDING).doubleValue();
    }

    public int getNumOfColumns() {
        return numOfColumns;
    }

    public Tbl getTable() {
        return tblBuilder.getObject();
    }

    /**
     * @param columnsIndices
     * @param percent
     * @return reference to this
     * @throws ArrayIndexOutOfBoundsException
     */
    public TableAdapter setColumnsWidth(int[] columnsIndices, double percent)
            throws ArrayIndexOutOfBoundsException {
        for (int i = 0; i < columnsIndices.length; i++) {
            setColumnWidth(columnsIndices[i], percent);
        }
        return this;
    }

    /**
     * @param columnsIndices
     * @param percents
     * @return reference to this
     * @throws ArrayIndexOutOfBoundsException
     */
    public TableAdapter setColumnsWidth(int[] columnsIndices, double[] percents)
            throws ArrayIndexOutOfBoundsException {
        for (int i = 0; i < columnsIndices.length; i++) {
            setColumnWidth(columnsIndices[i], percents[i]);
        }
        return this;
    }

    /**
     * @param columnIndex index of column
     * @param percent     percentage of column
     * @throws ArrayIndexOutOfBoundsException if columnIndex is out of range
     */
    public TableAdapter setColumnWidth(int columnIndex, Double percent) throws ArrayIndexOutOfBoundsException {
        setColumnWidthInternal(columnIndex, percent);
        return this;
    }

    /**
     * @param columnIndex index of column
     * @param percent     percentage of column
     * @throws ArrayIndexOutOfBoundsException if columnIndex is out of range
     */
    private void setColumnWidthInternal(int columnIndex, Double percent) throws ArrayIndexOutOfBoundsException {
        checkColumnIndex(columnIndex);
        BigDecimal bigPercent = new BigDecimal((percent == null) ? defaultPercent : percent);
        int width = TOTAL_TABLE_WIDTH.multiply(bigPercent, ROUNDING).divide(PERCENT).intValue();
        columnsWidths[columnIndex] = width;
        width = TOTAL_GRID_COL_WIDTH.multiply(bigPercent, ROUNDING).divide(PERCENT).intValue();
        gridWidths[columnIndex] = width;
    }

    /**
     * @return
     */
    public TableAdapter startRow() {
        return startRow(getTrBuilder().withRsidR(nextId()).withRsidTr(nextId()));
    }

    /**
     * @param trBuilder
     * @return
     */
    public TableAdapter startRow(TrBuilder trBuilder) {
        this.trBuilder = trBuilder;
        return this;
    }

    public TableAdapter startTable() {
        return startTable("TableGrid");
    }

    public TableAdapter startTable(String tableStyle) {
        return startTable(null, tableStyle);
    }

    public TableAdapter startTable(TblPr extraTblPr) {
        return startTable(extraTblPr, null);
    }

    public TableAdapter startTable(TblPr extraTblPr, String tableStyle) {
        TblGridBuilder tblGridBuilder = getTblGridBuilder();
        for (int i = 0; i < numOfColumns; i++) {
            tblGridBuilder.addGridCol(getTblGridColBuilder().withW(gridWidths[i].toString()).getObject());
        }

        TblWidth tblWidth = getTblWidthBuilder().withType(TYPE_PCT).withW(TOTAL_TABLE_WIDTH.toString()).getObject();

        CTTblLook cTTblLook = getCTTblLookBuilder().withFirstRow(ONE).withLastRow(ONE).withFirstColumn(ONE)
                .withLastColumn(ONE).withNoVBand(ONE).withNoHBand(ONE).getObject();

        tableStyle = isBlank(tableStyle) ? "TableGrid" : tableStyle;
        TblPr tblPr = getTblPrBuilder().withTblStyle(tableStyle).withTblW(tblWidth).withTblLook(cTTblLook).getObject();
        TblPrBuilder tblPrBuilder = new TblPrBuilder(tblPr, extraTblPr);
        tblPr = tblPrBuilder.getObject();

        tblBuilder.withTblGrid(tblGridBuilder.getObject()).withTblPr(tblPr);
        return this;
    }
}
