package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.util.AppUtil;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static org.testng.Assert.fail;

/**
 * @author sali
 */
public abstract class CommonTest {

    protected static Path docsPath;

    static {
        docsPath = Paths.get(System.getProperty("docs.dir", AppUtil.USER_TEMP_DIR.getAbsolutePath()));
        System.out.println(">>>>>>>>>>>> " + docsPath);
        if (!Files.exists(docsPath)) {
            try {
                docsPath = Files.createDirectory(docsPath);
            } catch (IOException e) {
                throw new UncheckedIOException(e.getMessage(), e);
            }
        }
    }

    protected WordprocessingMLPackage wmlPackage;
    private MainDocumentPart mainDocumentPart;

    protected abstract String getFileName();

    protected abstract WordprocessingMLPackage loadWmlPackage() throws Docx4JException;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = loadWmlPackage();
            mainDocumentPart = wmlPackage.getMainDocumentPart();
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            WmlAdapter.save(Paths.get(docsPath.toString(), getFileName()).toFile(), wmlPackage);
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    protected final MainDocumentPart getMainDocumentPart() {
        return mainDocumentPart;
    }
}
