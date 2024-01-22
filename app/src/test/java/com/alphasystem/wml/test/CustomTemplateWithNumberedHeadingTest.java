package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.docx4j.builder.wml.WmlPackageBuilder;
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
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().withTemplatePath("META-INF/Custom.dotx");
        return new WmlPackageBuilder(inputs).multiLevelHeading().getPackage();
    }

}
