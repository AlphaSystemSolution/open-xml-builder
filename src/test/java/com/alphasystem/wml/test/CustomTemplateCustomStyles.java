package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.TableAdapter;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.testng.annotations.Test;

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
        TableAdapter tableAdapter = new TableAdapter(20.0, 80.0).startTable("AdmonitionTable");
        tableAdapter.startRow().addColumn(0, getParagraphWithStyle("DefaultTitle", "Important"))
                .addColumn(1, getParagraphWithStyle(null, "Some text")).endRow();
        getMainDocumentPart().addObject(tableAdapter.getTable());
    }

}
