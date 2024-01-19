package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class CustomTemplateTest extends TemplateTest {

    @Override
    protected String getFileName() {
        return "custom-template.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        try {
            return WmlPackageBuilder.createPackage("META-INF/Custom.dotx").getPackage();
        } catch (InvalidFormatException e) {
            throw new SystemException(e.getMessage(), e);
        }
    }

}
