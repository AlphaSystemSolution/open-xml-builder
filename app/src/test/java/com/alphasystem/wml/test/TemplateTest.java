package com.alphasystem.wml.test;

import com.alphasystem.commons.util.AppUtil;
import com.alphasystem.docx4j.builder.wml.*;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.STAlignH;
import org.docx4j.dml.wordprocessingDrawing.STAlignV;
import org.docx4j.dml.wordprocessingDrawing.STRelFromH;
import org.docx4j.dml.wordprocessingDrawing.STRelFromV;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.testng.annotations.Test;

import java.net.URI;

import static com.alphasystem.docx4j.builder.wml.WmlAdapter.*;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRBuilder;
import static org.testng.Assert.assertEquals;
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
        final var p = getPBuilder().addContent(hyperlink).getObject();
        getMainDocumentPart().addObject(p);
    }

    @Test(dependsOnMethods = {"testHeading5"})
    public void addTableOfContent() {
        new TocGenerator().mainDocumentPart(getMainDocumentPart()).level(5).generateToc();
    }

    @Test(dependsOnMethods = {"addTableOfContent"})
    public void insertImage() throws Exception {
        final var url = AppUtil.getResource("java_logo.png");
        final var imageRun = new ImageAdapter(wmlPackage).withUrl(url).insertImage();
        getMainDocumentPart().addObject(getImagePara("Insert image with default size and no hints", imageRun));
        assertTrue(wmlPackage.getParts().getParts().values().stream()
                .anyMatch(part -> part instanceof BinaryPartAbstractImage), "Image must be embedded in the package");
        assertTrue(getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship().stream()
                .filter(rel -> rel.getType().endsWith("/image"))
                .noneMatch(rel -> "External".equals(rel.getTargetMode())), "Image must not use an external link");
    }

    @Test(dependsOnMethods = {"insertImage"})
    public void insertImageFromUrl() throws Exception {
        final var url = URI.create("https://www.vectorlogo.zone/logos/java/java-icon.svg").toURL();
        final var imageRun = new ImageAdapter(wmlPackage).withUrl(url).withFilenameHint("java-icon.svg").withAltText("java logo")
                .insertImage();
        getMainDocumentPart().addObject(getImagePara("Insert image from URL with default size", imageRun));
    }

    @Test(dependsOnMethods = {"insertImageFromUrl"})
    public void insertImageWithWidth() throws Exception {
        final var url = AppUtil.getResource("java_logo.png");
        final var imageRun = new ImageAdapter(wmlPackage).withUrl(url).withWidth(20.0).insertImage();
        getMainDocumentPart().addObject(getImagePara("Insert image with width only", imageRun));
    }

    @Test(dependsOnMethods = {"insertImageWithWidth"})
    public void insertImageWithWidthAndHeight() throws Exception {
        final var url = AppUtil.getResource("java_logo.png");
        final var imageRun = new ImageAdapter(wmlPackage).withUrl(url).withWidth(200.0).withHeight(100.0).insertImage();
        getMainDocumentPart().addObject(getImagePara("Insert image with width and height", imageRun));
    }

    @Test(dependsOnMethods = {"insertImageWithWidthAndHeight"})
    public void insertAnchoredImage() throws Exception {
        final var url = AppUtil.getResource("java_logo.png");
        final var imageRun = new ImageAdapter(wmlPackage).withUrl(url)
                .withHorizontalAlignment(STRelFromH.COLUMN, STAlignH.RIGHT)
                .withVerticalAlignment(STRelFromV.LINE, STAlignV.CENTER)
                .insertImage();
        getMainDocumentPart().addObject(getImagePara("Insert image anchored right/center of line", imageRun));

        final var drawing = imageRun.getContent().stream()
                .filter(Drawing.class::isInstance)
                .map(Drawing.class::cast)
                .findFirst()
                .orElseThrow(() -> new AssertionError("No drawing found in image run"));
        final var anchor = drawing.getAnchorOrInline().stream()
                .filter(Anchor.class::isInstance)
                .map(Anchor.class::cast)
                .findFirst()
                .orElseThrow(() -> new AssertionError("Drawing must contain an anchor"));
        assertEquals(anchor.getPositionH().getRelativeFrom(), STRelFromH.COLUMN);
        assertEquals(anchor.getPositionH().getAlign(), STAlignH.RIGHT);
        assertEquals(anchor.getPositionV().getRelativeFrom(), STRelFromV.LINE);
        assertEquals(anchor.getPositionV().getAlign(), STAlignV.CENTER);
    }

    private P getImagePara(String captionText, R imageRun) {
        final var textRun = getRBuilder().addContent(getText(captionText)).getObject();
        return getPBuilder().addContent(imageRun, getLineBreakRun(), textRun).getObject();
    }
}
