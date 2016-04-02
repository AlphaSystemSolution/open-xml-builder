package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

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

}
