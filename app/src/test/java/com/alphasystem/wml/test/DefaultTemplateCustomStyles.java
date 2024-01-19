package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
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
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useDefaultTemplate();
        return new WmlPackageBuilder(inputs).styles("META-INF/custom-styles.xml").getPackage();
    }

}
