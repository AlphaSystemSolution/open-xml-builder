package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.docx4j.builder.wml.WmlAdapter;
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
        docsPath = Paths.get(System.getProperty("docs.dir"));
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

    protected abstract WordprocessingMLPackage loadWmlPackage() throws SystemException;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = loadWmlPackage();
            mainDocumentPart = wmlPackage.getMainDocumentPart();
        } catch (SystemException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            final var file = Paths.get(docsPath.toString(), getFileName()).toFile();
            WmlAdapter.save(file, wmlPackage);
        } catch (Throwable e) {
            fail(e.getMessage(), e);
        }
    }

    protected final MainDocumentPart getMainDocumentPart() {
        return mainDocumentPart;
    }
}
