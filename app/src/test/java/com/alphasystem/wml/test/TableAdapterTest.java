package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.WmlBuilderFactory;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.ColumnData;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import com.alphasystem.openxml.builder.wml.table.TableType;
import com.alphasystem.openxml.builder.wml.table.VerticalMergeType;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRBuilder;

public class TableAdapterTest extends CommonTest {

    @Override
    protected String getFileName() {
        return "table-adapter-test.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        return new WmlPackageBuilder().styles("META-INF/custom-styles.xml").getPackage();
    }

    @Test
    public void testTableAdapter() {
        var mainDocumentPart = getMainDocumentPart();
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "PCT Table");
        var tableAdapter = new TableAdapter().withWidths(25.0, 8.0, 17.0, 17.0, 8.0, 25.0).startTable();

        tableAdapter.startRow()
                .addColumn(new ColumnData(0).withGridSpanValue(6)
                        .withContent(WmlAdapter.getParagraph("Column spans all grid spans.")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withContent(WmlAdapter.getParagraph("1")))
                .addColumn(new ColumnData(1).withGridSpanValue(2).withContent(WmlAdapter.getParagraph("2")))
                .addColumn(new ColumnData(3).withGridSpanValue(2).withContent(WmlAdapter.getParagraph("3")))
                .addColumn(new ColumnData(5).withContent(WmlAdapter.getParagraph("5")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withGridSpanValue(2)
                        .withContent(WmlAdapter.getParagraph("Column 1 of row with 3 columns")))
                .addColumn(new ColumnData(2).withGridSpanValue(2)
                        .withContent(WmlAdapter.getParagraph("Column 2 of row with 3 columns")))
                .addColumn(new ColumnData(4).withGridSpanValue(2)
                        .withContent(WmlAdapter.getParagraph("Column 3 of row with 3 columns")))
                .endRow();

        mainDocumentPart.addObject(tableAdapter.getTable());
        mainDocumentPart.addObject(WmlAdapter.getEmptyPara());

        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "PCT table with five columns");
        tableAdapter = new TableAdapter().withNumOfColumns(5).startTable();
        addColumns(tableAdapter);
        mainDocumentPart.addObject(tableAdapter.getTable());
        mainDocumentPart.addObject(WmlAdapter.getEmptyPara());

        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Auto Table");
        tableAdapter = new TableAdapter().withTableType(TableType.AUTO).withNumOfColumns(5).startTable();
        addColumns(tableAdapter);
        mainDocumentPart.addObject(tableAdapter.getTable());

        mainDocumentPart.addObject(WmlAdapter.getEmptyPara());
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Table within listing");

        mainDocumentPart.addObject(createNumberedParagraph("List 1"));

        tableAdapter = new TableAdapter()
                .withTableType(TableType.AUTO)
                .withNumOfColumns(5)
                .withIndentLevel(0)
                .startTable();
        addColumns(tableAdapter);
        mainDocumentPart.addObject(tableAdapter.getTable());
        mainDocumentPart.addObject(createNumberedParagraph("List 2"));
        var tblBorders = WmlBuilderFactory
                .getTblBordersBuilder()
                .withTop(WmlAdapter.getNilBorder())
                .withBottom(WmlAdapter.getNilBorder())
                .withLeft(WmlAdapter.getNilBorder())
                .withRight(WmlAdapter.getNilBorder())
                .withInsideH(WmlAdapter.getNilBorder())
                .withInsideV(WmlAdapter.getNilBorder())
                .getObject();
        var tblPr = WmlBuilderFactory
                .getTblPrBuilder()
                .withTblBorders(tblBorders)
                .getObject();
        tableAdapter = new TableAdapter()
                .withTableType(TableType.AUTO)
                .withIndentLevel(0)
                .withWidths(32.0, 17.0, 17.0, 17.0, 17.0)
                .withTableProperties(tblPr)
                .startTable();
        addColumns(tableAdapter);
        mainDocumentPart.addObject(tableAdapter.getTable());
    }

    @Test(dependsOnMethods = {"testTableAdapter"})
    public void testRowSpan() {
        var mainDocumentPart = getMainDocumentPart();
        mainDocumentPart.addObject(WmlAdapter.getEmptyPara());
        var tableAdapter = new TableAdapter().withNumOfColumns(4).startTable()
                .startRow()
                .addColumn(new ColumnData(0).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 1 and 2 Column1")))
                .addColumn(new ColumnData(1).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 1, 2, 3, and 4 Column 2")))
                .addColumn(new ColumnData(2).withContent(WmlAdapter.getParagraph("Row 1 Column 3")))
                .addColumn(new ColumnData(3).withContent(WmlAdapter.getParagraph("Row1 Column 4")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(1).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(2).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 2 and 3 Column 3")))
                .addColumn(new ColumnData(3).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 2 Column 4")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 3 and 4 Column 1")))
                .addColumn(new ColumnData(1).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(2).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(3).withVerticalMergeType(VerticalMergeType.RESTART)
                        .withContent(WmlAdapter.getParagraph("Row 3 and 4 Column 4")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(1).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(2).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .addColumn(new ColumnData(3).withVerticalMergeType(VerticalMergeType.CONTINUE)
                        .withContent(WmlAdapter.getEmptyPara()))
                .endRow();

        mainDocumentPart.addObject(tableAdapter.getTable());
    }

    private P createNumberedParagraph(String text) {
        var ppr = WmlAdapter.getListParagraphProperties(1L, 0L, true);
        var run = getRBuilder().addContent(getText(text)).getObject();
        return getPBuilder().withPPr(ppr).addContent(run).getObject();
    }

    private void addColumns(TableAdapter tableAdapter) {
        tableAdapter.startRow()
                .addColumn(new ColumnData(0).withContent(WmlAdapter.getParagraph("Row 1 Column 1")))
                .addColumn(new ColumnData(1).withContent(WmlAdapter.getParagraph("Row 1 Column 2")))
                .addColumn(new ColumnData(2).withContent(WmlAdapter.getParagraph("Row 1 Column 3")))
                .addColumn(new ColumnData(3).withContent(WmlAdapter.getParagraph("Row 1 Column 4")))
                .addColumn(new ColumnData(4).withContent(WmlAdapter.getParagraph("Row 1 Column 5")))
                .endRow();
    }
}
