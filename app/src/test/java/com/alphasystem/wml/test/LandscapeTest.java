package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * @author sali
 */
public class LandscapeTest extends TemplateTest  {

    @Override
    protected String getFileName() {
        return "landscape.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        try {
            return WmlPackageBuilder.createPackage(true).getPackage();
        } catch (InvalidFormatException e) {
            throw new SystemException(e.getMessage(), e);
        }
    }
}
