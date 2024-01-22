package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class DefaultTemplateWithNumberedHeadingTest extends TemplateTest {

    @Override
    protected String getFileName() {
        return "default-template-numbered-heading.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useDefaultTemplate();
        return new WmlPackageBuilder(inputs).multiLevelHeading().getPackage();
    }

//    @Test
//    public void addTableOfContents() {
//        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
//        final List<P> list = addTableOfContent();
//        list.forEach(mainDocumentPart::addObject);
//    }
//
//    @Test(dependsOnMethods = {"addTableOfContents"})
//    public void createMultiLevelHeading1() {
//        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
//        mainDocumentPart.addObject(getEmptyPara());
//
//        for (int i = 1; i <= 5; i++) {
//            String style = format("ListHeading%s", i);
//            mainDocumentPart.addStyledParagraphOfText(style, style);
//        }
//        mainDocumentPart.addObject(getPageBreak());
//    }
//
//    @Test(dependsOnMethods = {"createMultiLevelHeading1"})
//    public void createMultiLevelHeading2() {
//        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
//        mainDocumentPart.addObject(getEmptyPara());
//
//        for (int i = 1; i <= 5; i++) {
//            String style = format("ListHeading%s", i);
//            mainDocumentPart.addStyledParagraphOfText(style, style);
//        }
//    }

}
