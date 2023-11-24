package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
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
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return WmlPackageBuilder.createPackage("META-INF/Custom.dotx")
                .styles("META-INF/custom-styles.xml")
                .multiLevelHeading(EXAMPLE)
                .getPackage();
    }

    @Test
    public void testCustomTableStyle() {
        TableAdapter tableAdapter = new TableAdapter().startTable("AdmonitionTable", 10.0, 90.0);
        tableAdapter.startRow().addColumn(0, getParagraph("Important"))
                .addColumn(1, getParagraph("Some text")).endRow();
        getMainDocumentPart().addObject(tableAdapter.getTable());
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateHorizontalList() {
        TableAdapter tableAdapter = new TableAdapter().startTable("HorizontalList",15.0, 85.0);
        tableAdapter.startRow().addColumn(0, getParagraph("Some Long Title"))
                .addColumn(1, getParagraph("Some text")).endRow();
        getMainDocumentPart().addObject(tableAdapter.getTable());
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithHeader(){
        TableAdapter tableAdapter = new TableAdapter().startTable("TableGrid1",33.0, 33.0, 33.0);
        tableAdapter.startRow().addColumn(0, getParagraph("Header 1")).addColumn(1, getParagraph("Header 2"))
                .addColumn(2, getParagraph("Header 3")).endRow();
        tableAdapter.startRow().addColumn(0, getParagraph("Column 1 Row 1")).addColumn(1, getParagraph("Column 2 Row 1"))
                .addColumn(2, getParagraph("Column 3 Row 1")).endRow();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Header Row"));
        getMainDocumentPart().addObject(tableAdapter.getTable());
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithFooter(){
        TableAdapter tableAdapter = new TableAdapter().startTable("TableGrid2",33.0, 33.0, 33.0);
        tableAdapter.startRow().addColumn(0, getParagraph("Column 1 Row 1")).addColumn(1, getParagraph("Column 2 Row 1"))
                .addColumn(2, getParagraph("Column 3 Row 1")).endRow();
        tableAdapter.startRow().addColumn(0, getParagraph("Footer 1")).addColumn(1, getParagraph("Footer 2"))
                .addColumn(2, getParagraph("Footer 3")).endRow();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Footer Row"));
        getMainDocumentPart().addObject(tableAdapter.getTable());
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test
    public void testCreateTableWithHeaderAndFooter(){
        TableAdapter tableAdapter = new TableAdapter().startTable("TableGrid3",33.0, 33.0, 33.0);
        tableAdapter.startRow().addColumn(0, getParagraph("Header 1")).addColumn(1, getParagraph("Header 2"))
                .addColumn(2, getParagraph("Header 3")).endRow();
        tableAdapter.startRow().addColumn(0, getParagraph("Column 1 Row 1")).addColumn(1, getParagraph("Column 2 Row 1"))
                .addColumn(2, getParagraph("Column 3 Row 1")).endRow();
        tableAdapter.startRow().addColumn(0, getParagraph("Footer 1")).addColumn(1, getParagraph("Footer 2"))
                .addColumn(2, getParagraph("Footer 3")).endRow();
        getMainDocumentPart().addObject(getParagraphWithStyle("DefaultTitle", "Table with Header & Footer Row"));
        getMainDocumentPart().addObject(tableAdapter.getTable());
        getMainDocumentPart().addObject(getEmptyPara());
    }

}
