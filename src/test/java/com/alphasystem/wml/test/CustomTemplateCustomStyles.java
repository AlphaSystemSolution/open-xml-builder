package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.TableAdapter;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getParagraph;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getParagraphWithStyle;
import static com.alphasystem.wml.test.DocumentCaption.EXAMPLE;

/**
 * @author sali
 */
public class CustomTemplateCustomStyles extends CustomStylesTest {

    @Override
    protected String getFileName() {
        return "custom-template-custom-styles.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return new WmlPackageBuilder("META-INF/Custom.dotx").styles("META-INF/custom-styles.xml").multiLevelHeading(EXAMPLE).getPackage();
    }

    @Test
    public void testCustomTableStyle() {
        TableAdapter tableAdapter = new TableAdapter(10.0, 90.0).startTable("AdmonitionTable");
        tableAdapter.startRow().addColumn(0, getParagraphWithStyle("DefaultTitle", "Important"))
                .addColumn(1, getParagraph("Some text")).endRow();
        getMainDocumentPart().addObject(tableAdapter.getTable());
    }

    @Test
    public void testCreateHorizontalList() {
        TableAdapter tableAdapter = new TableAdapter(15.0, 85.0).startTable("HorizontalList");
        tableAdapter.startRow().addColumn(0, getParagraph("Some Long Title"))
                .addColumn(1, getParagraph("Some text")).endRow();
        getMainDocumentPart().addObject(tableAdapter.getTable());
    }

}
