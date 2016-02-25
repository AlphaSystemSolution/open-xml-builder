package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getEmptyPara;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.save;
import static java.lang.String.format;
import static java.nio.file.Paths.get;
import static org.testng.Assert.fail;

/**
 * @author sali
 */
public class MultiLevelHeadingTest {

    private static final String PARENT_PATH = "C:\\Users\\sali\\git-hub\\AlphaSystemSolution\\open-xml-builder\\target";

    private WordprocessingMLPackage wmlPackage;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = new WmlPackageBuilder().multiLevelHeading().getPackage();
        } catch (InvalidFormatException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            save(get(PARENT_PATH, "multi-level-heading.docx").toFile(), wmlPackage);
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }


    @Test
    public void createMultiLevelHeading1() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        for (int i = 1; i <= 5; i++) {
            String style = format("Heading%s", i);
            mainDocumentPart.addStyledParagraphOfText(style, style);
        }
    }

    @Test
    public void createMultiLevelHeading2() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        for (int i = 1; i <= 5; i++) {
            String style = format("Heading%s", i);
            mainDocumentPart.addStyledParagraphOfText(style, style);
        }
    }

}
