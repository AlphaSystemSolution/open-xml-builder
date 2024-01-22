package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class DefaultTemplateTest extends TemplateTest {

    @Override
    protected String getFileName() {
        return "default-template.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useDefaultTemplate().useDefaultStyles();
        return new WmlPackageBuilder(inputs).getPackage();
    }

}
