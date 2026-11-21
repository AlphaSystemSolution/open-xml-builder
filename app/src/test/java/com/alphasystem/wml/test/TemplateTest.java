package com.alphasystem.wml.test;

import com.alphasystem.commons.util.AppUtil;
import com.alphasystem.docx4j.builder.wml.TocGenerator;
import com.alphasystem.docx4j.builder.wml.WmlAdapter;
import com.alphasystem.docx4j.builder.wml.WmlBuilderFactory;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.P;
import org.testng.annotations.Test;

import static com.alphasystem.docx4j.builder.wml.WmlAdapter.*;
import static org.testng.Assert.assertTrue;

/**
 * @author sali
 */
public abstract class TemplateTest extends CommonTest {

    @Test
    public void testTitle() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Title", "Document Title"));
        getMainDocumentPart().addObject(getHorizontalLine());
    }

    @Test(dependsOnMethods = {"testTitle"})
    public void testHeading1() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading1", "Heading Level 1"));
        getMainDocumentPart().addObject(getParagraphWithStyle(null, "A paragraph with document default font"));
    }

    @Test(dependsOnMethods = {"testHeading1"})
    public void testHeading2() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading2", "Heading Level 2"));
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test(dependsOnMethods = {"testHeading2"})
    public void testHeading3() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading3", "Heading Level 3"));
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test(dependsOnMethods = {"testHeading3"})
    public void testHeading4() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading4", "Heading Level 4"));
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test(dependsOnMethods = {"testHeading4"})
    public void testHeading5() {
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading5", "Heading Level 5"));
        getMainDocumentPart().addObject(getEmptyPara());
    }

    @Test(dependsOnMethods = {"testHeading5"})
    public void testAnotherHeading1() {
        getMainDocumentPart().addObject(getPageBreakPara());
        getMainDocumentPart().addObject(getParagraphWithStyle("Heading1", "Another Heading Level 1"));
        getMainDocumentPart().addObject(getEmptyPara());
        final P.Hyperlink hyperlink = addHyperlink("contents", "Back to top");
        final var p = WmlBuilderFactory.getPBuilder().addContent(hyperlink).getObject();
        getMainDocumentPart().addObject(p);
    }

    @Test(dependsOnMethods = {"testHeading5"})
    public void addTableOfContent() {
        new TocGenerator().mainDocumentPart(getMainDocumentPart()).level(5).generateToc();
    }

    @Test(dependsOnMethods = {"addTableOfContent"})
    public void insertImage() throws Exception {
        final var url = AppUtil.getResource("java_logo.png");
        final var image = WmlAdapter.insertImage(wmlPackage, "java_logo.png", "Java Logo", 1L, 1, url);
        getMainDocumentPart().addObject(image);
        assertTrue(wmlPackage.getParts().getParts().values().stream()
                .anyMatch(part -> part instanceof BinaryPartAbstractImage), "Image must be embedded in the package");
        assertTrue(getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship().stream()
                .filter(rel -> rel.getType().endsWith("/image"))
                .noneMatch(rel -> "External".equals(rel.getTargetMode())), "Image must not use an external link");
    }
}
