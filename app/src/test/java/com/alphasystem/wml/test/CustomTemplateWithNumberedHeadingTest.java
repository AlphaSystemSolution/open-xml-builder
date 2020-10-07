package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class CustomTemplateWithNumberedHeadingTest extends TemplateTest {

    @Override
    protected String getFileName() {
        return "custom-template-numbered-heading.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return WmlPackageBuilder.createPackage("META-INF/Custom.dotx").multiLevelHeading().getPackage();
    }

}
