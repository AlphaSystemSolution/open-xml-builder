package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.WmlBuilderFactory;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import com.alphasystem.openxml.builder.wml.table.TableType;
import org.docx4j.openpackaging.exceptions.Docx4JException;
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
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return WmlPackageBuilder.createPackage().styles("META-INF/custom-styles.xml").getPackage();
    }

    @Test
    public void testTableAdapter() {
        var mainDocumentPart = getMainDocumentPart();
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "PCT Table");
        var tableAdapter = new TableAdapter().withWidths(25.0, 8.0, 17.0, 17.0, 8.0, 25.0).startTable();

        tableAdapter.startRow()
                .addColumn(0, 6, WmlAdapter.getParagraph("Column spans all grid spans."))
                .endRow()
                .startRow()
                .addColumn(0, 1, WmlAdapter.getParagraph("1"))
                .addColumn(1, 2, WmlAdapter.getParagraph("2"))
                .addColumn(3, 2, WmlAdapter.getParagraph("3"))
                .addColumn(5, 1, WmlAdapter.getParagraph("4"))
                .endRow()
                .startRow()
                .addColumn(0, 2, WmlAdapter.getParagraph("Column 1 of row with 3 columns"))
                .addColumn(2, 2, WmlAdapter.getParagraph("Column 2 of row with 3 columns"))
                .addColumn(4, 2, WmlAdapter.getParagraph("Column 3 of row with 3 columns"))
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

    private P createNumberedParagraph(String text) {
        var ppr = WmlAdapter.getListParagraphProperties(1L, 0L, true);
        var run = getRBuilder().addContent(getText(text)).getObject();
        return getPBuilder().withPPr(ppr).addContent(run).getObject();
    }

    private void addColumns(TableAdapter tableAdapter) {
        tableAdapter.startRow()
                .addColumn(0, 1, WmlAdapter.getParagraph("Row 1 Column 1"))
                .addColumn(1, 1, WmlAdapter.getParagraph("Row 1 Column 2"))
                .addColumn(2, 1, WmlAdapter.getParagraph("Row 1 Column 3"))
                .addColumn(3, 1, WmlAdapter.getParagraph("Row 1 Column 4"))
                .addColumn(4, 1, WmlAdapter.getParagraph("Row 1 Column 5"))
                .endRow();
    }
}
