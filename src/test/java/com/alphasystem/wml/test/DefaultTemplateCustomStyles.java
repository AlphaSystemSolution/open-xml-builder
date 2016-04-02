package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class DefaultTemplateCustomStyles extends CustomStylesTest {

    @Override
    protected String getFileName() {
        return "default-template-custom-styles.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return new WmlPackageBuilder().styles("META-INF/custom-styles.xml").getPackage();
    }

}
