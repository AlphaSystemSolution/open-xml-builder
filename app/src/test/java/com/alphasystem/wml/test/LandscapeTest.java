package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.docx4j.builder.wml.WmlPackageBuilder;
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
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useLandscape();
        return new WmlPackageBuilder(inputs).getPackage();
    }
}
