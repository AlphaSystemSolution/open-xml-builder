package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

import java.io.File;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static java.lang.System.getProperty;
import static org.testng.Assert.fail;

/**
 * @author sali
 */
public abstract class CommonTest {

    private static final String USER_DIR = getProperty("user.dir", ".");
    private static final String USER_HOME = getProperty("user.home", USER_DIR);
    private static final File USER_TEMP_DIR = new File(getProperty("java.io.tmpdir", USER_HOME));

    protected static Path docsPath;

    static {
        docsPath = Paths.get(System.getProperty("docs.dir", USER_TEMP_DIR.getAbsolutePath()));
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
            final File file = Paths.get(docsPath.toString(), getFileName()).toFile();
            WmlAdapter.save(file, wmlPackage);
            WmlAdapter.saveAsPdf(file, wmlPackage);
        } catch (Exception e) {
            fail(e.getMessage(), e);
        }
    }

    protected final MainDocumentPart getMainDocumentPart() {
        return mainDocumentPart;
    }
}
