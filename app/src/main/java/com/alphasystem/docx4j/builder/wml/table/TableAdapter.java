package com.alphasystem.docx4j.builder.wml.table;

import com.alphasystem.docx4j.builder.wml.*;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.wml.*;

import java.math.BigDecimal;
import java.math.MathContext;
import java.math.RoundingMode;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.stream.IntStream;

import static com.alphasystem.commons.util.IdGenerator.nextId;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;
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
    public static final MathContext ROUNDING = new MathContext(6, RoundingMode.FLOOR);

    private TableType tableType;
    private String tableStyle;
    private int indentLevel;
    private ColumnInput[] inputs;
    private TblPr tableProperties;
    private ColumnAdapter columnAdapter;
    private TblBuilder tblBuilder;
    private TrBuilder trBuilder;

    public TableAdapter() {
        this.tableType = TableType.PCT;
        this.tableStyle = DEFAULT_TABLE_STYLE;
        this.indentLevel = -1;
        this.inputs = null;
    }

    public TableAdapter withTableType(TableType tableType) {
        this.tableType = tableType == null ? TableType.PCT : tableType;
        return this;
    }

    public TableAdapter withTableStyle(String tableStyle) {
        this.tableStyle = StringUtils.isBlank(tableStyle) ? DEFAULT_TABLE_STYLE : tableStyle;
        return this;
    }

    public TableAdapter withIndentLevel(int indentLevel) {
        this.indentLevel = indentLevel;
        return this;
    }

    public TableAdapter withNumOfColumns(int numOfColumns) {
        require(numOfColumns > 0, "`numOfColumns` must be positive integer");

        var divisor = BigDecimal.valueOf(numOfColumns);

        // each column should be of same size
        var columnWidth = PERCENT.divide(divisor, ROUNDING);
        var totalCalculatedWidth = columnWidth.multiply(divisor);
        var diff = PERCENT.subtract(totalCalculatedWidth);

        var columnWidths = new BigDecimal[numOfColumns];
        Arrays.fill(columnWidths, columnWidth);
        // update last column with any difference between total width and calculated width
        columnWidths[columnWidths.length - 1] = columnWidth.add(diff);

        return withWidths(Arrays.stream(columnWidths).map(BigDecimal::doubleValue).toArray(Double[]::new));
    }

    public TableAdapter withWidths(Double... widths) {
        var length = Objects.requireNonNull(widths, "Parameter 'widths' cannot be null").length;
        var inputs = IntStream.range(0, length)
                .mapToObj(i -> new ColumnInput(String.format("col_%s", i + 1), widths[i]))
                .toArray(ColumnInput[]::new);
        return withColumnInputs(inputs);
    }

    public TableAdapter withColumnInputs(ColumnInput... inputs) {
        var length = Objects.requireNonNull(inputs, "Parameter 'inputs' cannot be null").length;
        this.inputs = length == 0 ? new ColumnInput[]{new ColumnInput("col_1", PERCENT.doubleValue())} : inputs;

        // validate sum of all widths are equal to 0
        var sum = Arrays.stream(this.inputs).map(ColumnInput::getColumnWidth).map(BigDecimal::valueOf)
                .reduce(BigDecimal.ZERO, (sumValue, current) -> sumValue.add(current, ROUNDING));
        var diff = PERCENT.subtract(sum).doubleValue();
        require(diff == 0.0, String.format("Total column widths must be equal to 100 instead got %s", sum));

        return this;
    }

    public TableAdapter withTableProperties(TblPr tableProperties) {
        this.tableProperties = tableProperties;
        return this;
    }

    public TableAdapter startTable() {
        tblBuilder = getTblBuilder();
        this.columnAdapter = new ColumnAdapter(tableType, indentLevel, inputs);

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
        final var tblPrBuilder = new TblPrBuilder(tblPr, tableProperties);

        tblBuilder.withTblGrid(tblGridBuilder.getObject()).withTblPr(tblPrBuilder.getObject());
        return this;
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

    public TableAdapter addColumn(ColumnData columnData) {
        trBuilder.addContent(createColumn(tableType, columnData, getColumns()));
        return this;
    }

    public Tbl getTable() {
        return tblBuilder.getObject();
    }

    public List<ColumnInfo> getColumns() {
        return columnAdapter.getColumns();
    }

    public static Tc createColumn(TableType tableType, ColumnData columnData, List<ColumnInfo> columnInfos) {
        final var columnProperties = getColumnProperties(tableType, columnData.getColumnIndex(), columnData.getGridSpanValue(),
                columnData.getVerticalMergeType(), columnData.getColumnProperties(), columnInfos);
        return getTcBuilder().withTcPr(columnProperties).addContent(columnData.getContent()).getObject();
    }

    public static TcPr getColumnProperties(TableType tableType,
                                           Integer columnIndex,
                                           Integer gridSpanValue,
                                           VerticalMergeType verticalMergeType,
                                           TcPr columnProperties,
                                           List<ColumnInfo> columns)
            throws ArrayIndexOutOfBoundsException {
        checkColumnIndex(columns, columnIndex);
        final var columnInfo = columns.get(columnIndex);
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

        final var tcPrBuilder = getTcPrBuilder();

        final var tblWidth = getTblWidthBuilder().withType(tableType.getColumnType()).withW(columnWidth.longValue()).getObject();
        tcPrBuilder.withGridSpan(gs).withTcW(tblWidth);

        if (verticalMergeType != null && !VerticalMergeType.NONE.equals(verticalMergeType)) {
            final var vMerge = tcPrBuilder.getVMergeBuilder().withVal(verticalMergeType.getValue()).getObject();
            tcPrBuilder.withVMerge(vMerge);
        }

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

    private static void require(boolean condition, String message) {
        var prefix = "Requirement failed";
        var msg = StringUtils.isBlank(message) ? prefix : String.format("%s: %s", prefix, message);
        if (!condition) {
            throw new IllegalArgumentException(msg);
        }
    }
}

