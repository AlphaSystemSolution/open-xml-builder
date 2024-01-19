package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.ColumnData;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.*;
import static com.alphasystem.wml.test.DocumentCaption.EXAMPLE;

/**
 * @author sali
 */
public class CustomTemplateCustomStylesTest extends CustomStylesTest {

    @Override
    protected String getFileName() {
        return "custom-template-custom-styles.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        try {
            return WmlPackageBuilder.createPackage("META-INF/Custom.dotx")
                    .styles("META-INF/custom-styles.xml")
                    .multiLevelHeading(EXAMPLE)
                    .getPackage();
        } catch (InvalidFormatException e) {
            throw new SystemException(e.getMessage(), e);
        }
    }

    @Test
    public void testCustomTableStyle() {
        var table = new TableAdapter()
                .withTableStyle("AdmonitionTable")
                .withWidths(10.0, 90.0)
                .startTable()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Important")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Some text")))
                .endRow()
                .getTable();
        getMainDocumentPart().addObject(table);
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateHorizontalList() {
        var table = new TableAdapter()
                .withTableStyle("HorizontalList")
                .withWidths(15.0, 85.0)
                .startTable()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Some Long Title")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Some text")))
                .endRow()
                .getTable();
        getMainDocumentPart().addObject(table);
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithHeader(){
        var table = new TableAdapter()
                .withTableStyle("TableGrid1")
                .withWidths(33.3333, 33.3333, 33.3334)
                .startTable()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Header 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Header 2")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Header 3")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Column 1 Row 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Column 2 Row 1")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Column 3 Row 1")))
                .endRow()
                .getTable();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Header Row"));
        getMainDocumentPart().addObject(table);
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithFooter(){
        var table = new TableAdapter()
                .withTableStyle("TableGrid2")
                .withWidths(33.3333, 33.3333, 33.3334)
                .startTable()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Column 1 Row 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Column 2 Row 1")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Column 3 Row 1")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Footer 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Footer 2")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Footer 3")))
                .endRow()
                .getTable();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Footer Row"));
        getMainDocumentPart().addObject(table);
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithHeaderAndFooter(){
        var table = new TableAdapter()
                .withTableStyle("TableGrid3")
                .withWidths(33.3333, 33.3333, 33.3334)
                .startTable()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Header 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Header 2")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Header 3")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Column 1 Row 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Column 2 Row 1")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Column 3 Row 1")))
                .endRow()
                .startRow()
                .addColumn(new ColumnData(0).withContent(getParagraph("Footer 1")))
                .addColumn(new ColumnData(1).withContent(getParagraph("Footer 2")))
                .addColumn(new ColumnData(2).withContent(getParagraph("Footer 3")))
                .endRow()
                .getTable();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Header & Footer Row"));
        getMainDocumentPart().addObject(table);
        getMainDocumentPart().addObject(getEmptyPara());
    }

}
