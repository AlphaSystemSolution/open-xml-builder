package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.File;
import java.io.IOException;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getEmptyPara;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.save;
import static java.lang.String.format;
import static java.nio.file.Paths.get;
import static org.testng.Assert.fail;

/**
 * @author sali
 */
public class CustomTemplateWithListHeadingTest {

    private static final String PARENT_PATH = "C:\\Users\\sali\\git-hub\\AlphaSystemSolution\\open-xml-builder\\target";

    private WordprocessingMLPackage wmlPackage;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = new WmlPackageBuilder("META-INF/Custom.dotx").multiLevelHeading().getPackage();
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            final File file = get(PARENT_PATH, "CustomWithListHeading.docx").toFile();
            save(file, wmlPackage);
            Thread thread = new Thread(() -> {
                try {
                    Desktop.getDesktop().open(file);
                } catch (IOException e) {
                    e.printStackTrace();
                    // ignore
                }
            });
            thread.start();
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    @Test
    public void createMultiLevelHeading() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        for (int i = 1; i <= 5; i++) {
            String style = format("ListHeading%s", i);
            mainDocumentPart.addStyledParagraphOfText(style, style);
        }
    }
}
