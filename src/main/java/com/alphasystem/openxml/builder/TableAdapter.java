/**
 *
 */
package com.alphasystem.openxml.builder;

import com.alphasystem.openxml.builder.wml.TblBuilder;
import com.alphasystem.openxml.builder.wml.TblGridBuilder;
import com.alphasystem.openxml.builder.wml.TcPrBuilder;
import com.alphasystem.openxml.builder.wml.TrBuilder;
import org.docx4j.wml.*;
import org.docx4j.wml.CTTblPrBase.TblStyle;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.docx4j.wml.TcPrInner.TcBorders;

import java.math.BigDecimal;
import java.math.MathContext;

import static com.alphasystem.openxml.builder.OpenXmlAdapter.getGridSpan;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.math.RoundingMode.CEILING;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.docx4j.sharedtypes.STOnOff.ONE;
import static org.docx4j.sharedtypes.STOnOff.ZERO;

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
     * @param columnWidthPercentages
     * @throws IllegalArgumentException
     */
    public TableAdapter(Double... columnWidthPercentages)
            throws IllegalArgumentException {
        if (isEmpty(columnWidthPercentages)) {
            throw new IllegalArgumentException(
                    "Columns width are not initailized");
        }
        tblBuilder = getTblBuilder();
        numOfColumns = columnWidthPercentages.length;
        this.bigNumOfColumn = new BigDecimal(numOfColumns);
        this.gridWidths = new Integer[numOfColumns];
        this.columnsWidths = new Integer[numOfColumns];
        this.defaultPercent = getDefaultPercent();

        for (int index = 0; index < columnWidthPercentages.length; index++) {
            Double percent = columnWidthPercentages[index];
            setColumnWidth(index, percent, true);
        }
    }

    /**
     * @param numOfColumns
     */
    public TableAdapter(int numOfColumns) {
        this(new Double[numOfColumns]);
    }

    /**
     * @param columnIndex
     * @param gridSpanValue
     * @param borders
     * @param paras
     * @return
     * @throws ArrayIndexOutOfBoundsException
     */
    public TableAdapter addColumn(Integer columnIndex, Integer gridSpanValue,
                                  TcBorders borders, P... paras)
            throws ArrayIndexOutOfBoundsException {
        checkColumnIndex(columnIndex);
        Integer cw = columnsWidths[columnIndex];
        GridSpan gridSpan = getGridSpan(1);
        if (gridSpanValue != null && gridSpanValue > 1) {
            // sanity check, make sure we are not going out of bound
            checkColumnIndex(columnIndex + gridSpanValue - 1);
            // iterate through width and get the total width for the gridspan
            for (int i = columnIndex + 1; i < gridSpanValue; i++) {
                cw += columnsWidths[i];
            }
            gridSpan = getGridSpan(gridSpanValue);
        }
        TblWidth tblWidth = getTblWidthBuilder().withType(TYPE_PCT)
                .withW(cw.toString()).getObject();
        TcPrBuilder tcPrBuilder = getTcPrBuilder().withGridSpan(gridSpan)
                .withTcW(tblWidth);
        if (borders != null) {
            tcPrBuilder.withTcBorders(borders);
        }
        Tc tc = getTcBuilder().withTcPr(tcPrBuilder.getObject())
                .addContent((Object[]) paras).getObject();
        trBuilder.addContent(tc);
        return this;
    }

    /**
     * @param columnIndex
     * @throws ArrayIndexOutOfBoundsException
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
     * @return
     */
    public TableAdapter endRow() {
        tblBuilder.addContent(trBuilder.getObject());
        trBuilder = null;
        return this;
    }

    private double getDefaultPercent() {
        BigDecimal width = TOTAL_TABLE_WIDTH.divide(bigNumOfColumn, ROUNDING);
        return width.divide(TOTAL_TABLE_WIDTH).multiply(PERCENT, ROUNDING)
                .doubleValue();
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
     * @return
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
     * @return
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
     * @param columnIndex
     * @param percent
     * @throws ArrayIndexOutOfBoundsException
     */
    public TableAdapter setColumnWidth(int columnIndex, Double percent)
            throws ArrayIndexOutOfBoundsException {
        setColumnWidth(columnIndex, percent, true);
        return this;
    }

    /**
     * @param columnIndex
     * @param percent
     * @param dummy
     * @throws ArrayIndexOutOfBoundsException
     */
    private void setColumnWidth(int columnIndex, Double percent, boolean dummy)
            throws ArrayIndexOutOfBoundsException {
        checkColumnIndex(columnIndex);
        BigDecimal bigPercent = new BigDecimal(percent == null ? defaultPercent
                : percent);
        int width = TOTAL_TABLE_WIDTH.multiply(bigPercent, ROUNDING)
                .divide(PERCENT).intValue();
        columnsWidths[columnIndex] = width;
        width = TOTAL_GRID_COL_WIDTH.multiply(bigPercent, ROUNDING)
                .divide(PERCENT).intValue();
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

    /**
     *
     */
    public TableAdapter startTable() {
        TblGridBuilder tblGridBuilder = getTblGridBuilder();
        for (int i = 0; i < numOfColumns; i++) {
            tblGridBuilder.addGridCol(getTblGridColBuilder().withW(
                    gridWidths[i].toString()).getObject());
        }

        TblStyle tblStyle = getCTTblPrBaseTblStyleBuilder()
                .withVal("TableGrid").getObject();

        TblWidth tblWidth = getTblWidthBuilder().withType(TYPE_PCT)
                .withW(TOTAL_TABLE_WIDTH.toString()).getObject();

        CTTblLook cTTblLook = getCTTblLookBuilder().withFirstRow(ONE)
                .withLastRow(ZERO).withFirstColumn(ONE).withLastColumn(ZERO)
                .withNoVBand(ONE).withNoHBand(ZERO).getObject();

        TblPr tblPr = getTblPrBuilder().withTblStyle(tblStyle)
                .withTblW(tblWidth).withTblLook(cTTblLook).getObject();

        tblBuilder.withTblGrid(tblGridBuilder.getObject()).withTblPr(tblPr);
        return this;
    }
}
