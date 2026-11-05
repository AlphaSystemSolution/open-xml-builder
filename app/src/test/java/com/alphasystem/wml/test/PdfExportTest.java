package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.docx4j.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import org.testng.SkipException;
import org.testng.annotations.Test;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getParagraphWithStyle;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getText;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.registerFont;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.saveAsPdf;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.toPdf;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRBuilder;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRFontsBuilder;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRPrBuilder;
import static org.testng.Assert.assertEquals;
import static org.testng.Assert.assertTrue;
import static org.testng.Assert.fail;

/**
 * @author sali
 */
public class PdfExportTest {

    private static final Path docsPath = Paths.get(System.getProperty("docs.dir", "build/docs"));
    private static final File SUBSTITUTE_FONT = new File("/Library/Fonts/Arial Unicode.ttf");
    private static final File ARABIC_FONT = new File(
            System.getProperty("user.home"), "Library/Fonts/KFGQPC Uthman Taha Naskh Regular.ttf");

    private static WordprocessingMLPackage newPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useDefaultTemplate().useDefaultStyles();
        return new WmlPackageBuilder(inputs).getPackage();
    }

    private static void checkFontsAvailable() {
        if (!SUBSTITUTE_FONT.isFile() || !ARABIC_FONT.isFile()) {
            throw new SkipException("Required test fonts are not installed on this system");
        }
    }

    @Test
    public void testToPdf() throws Exception {
        checkFontsAvailable();
        final var wmlPackage = newPackage();
        wmlPackage.getMainDocumentPart().addObject(getParagraphWithStyle("Title", "Document Title"));
        wmlPackage.getMainDocumentPart().addObject(getParagraphWithStyle(null, "نص عربي للاختبار"));

        // the default styles reference MS fonts which may not be installed;
        // alias them to a local font file, and register the Arabic font file too
        registerFont("Calibri", SUBSTITUTE_FONT);
        registerFont("Calibri Light", SUBSTITUTE_FONT);
        registerFont("Times New Roman", SUBSTITUTE_FONT);
        final var pdfFile = docsPath.resolve("pdf-export.pdf").toFile();
        pdfFile.getParentFile().mkdirs();
        try (final var os = Files.newOutputStream(pdfFile.toPath())) {
            toPdf(wmlPackage, os, ARABIC_FONT);
        }
        assertTrue(pdfFile.isFile());
        assertTrue(pdfFile.length() > 0);
        try (final var is = Files.newInputStream(pdfFile.toPath())) {
            final var header = new byte[4];
            assertEquals(is.read(header), 4);
            assertEquals(new String(header), "%PDF");
        }
    }

    @Test
    public void testMissingFont() throws Exception {
        final var wmlPackage = newPackage();
        final var rFonts = getRFontsBuilder()
                .withAscii("NonExistentFontXYZ")
                .withHAnsi("NonExistentFontXYZ")
                .withCs("NonExistentFontXYZ")
                .getObject();
        final var rPr = getRPrBuilder().withRFonts(rFonts).getObject();
        final var run = getRBuilder().withRPr(rPr).addContent(getText("text")).getObject();
        final P p = getPBuilder().addContent(run).getObject();
        wmlPackage.getMainDocumentPart().addObject(p);
        try {
            docsPath.toFile().mkdirs();
            saveAsPdf(docsPath.resolve("pdf-export-missing-font.pdf").toFile(), wmlPackage);
            fail("Expected Docx4JException for missing font");
        } catch (Docx4JException e) {
            assertTrue(e.getMessage().contains("NonExistentFontXYZ"), e.getMessage());
        }
    }
}
