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

import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getParagraphWithStyle;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getText;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.registerFont;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.registerFonts;
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
public class PdfExportTest extends CustomStylesTest {

    private static final String FILE_NAME_PREFIX = "pdf-export-test";
    private static final File SUBSTITUTE_FONT = new File("/Library/Fonts/Arial Unicode.ttf");
    private static final File ARABIC_FONT = new File(
            System.getProperty("user.home"), "Library/Fonts/KFGQPC Uthman Taha Naskh Regular.ttf");

    @Override
    protected String getFileName() {
        return String.format("%s.docx", FILE_NAME_PREFIX);
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().useDefaultTemplate();
        return new WmlPackageBuilder(inputs).styles("META-INF/custom-styles.xml").getPackage();
    }

    private static void checkFontsAvailable() {
        if (!SUBSTITUTE_FONT.isFile() || !ARABIC_FONT.isFile()) {
            throw new SkipException("Required test fonts are not installed on this system");
        }
    }

    @Test
    public void testToPdf() throws Exception {
        checkFontsAvailable();
        getMainDocumentPart().addObject(getParagraphWithStyle("Title", "Document Title"));
        getMainDocumentPart().addObject(getParagraphWithStyle("CustomArabic", "سلم"));
        getMainDocumentPart().addObject(getParagraphWithStyle("CustomArabic", "نَصَرَ يَنْصُرُ نَصْرًا نَاصِرٌ"));

        // the default styles reference MS fonts which may not be installed;
        // alias them to a local font file, and register the Arabic font file too
        registerFont("Calibri", SUBSTITUTE_FONT);
        registerFont("Calibri Light", SUBSTITUTE_FONT);
        registerFont("Times New Roman", SUBSTITUTE_FONT);
        final var pdfFile = docsPath.resolve(String.format("%s.pdf", FILE_NAME_PREFIX)).toFile();
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
    public void testSaveAsPdfWithFontsDirectory() throws Exception {
        checkFontsAvailable();
        getMainDocumentPart().addObject(getParagraphWithStyle("Title", "Document Title"));
        getMainDocumentPart().addObject(getParagraphWithStyle("CustomArabic", "سلم"));

        // the default styles reference MS fonts which may not be installed;
        // alias them to a local font file
        registerFont("Calibri", SUBSTITUTE_FONT);
        registerFont("Calibri Light", SUBSTITUTE_FONT);
        registerFont("Times New Roman", SUBSTITUTE_FONT);

        // build a fonts directory containing a real font plus non-font and
        // invalid-font files, which must be silently ignored
        final var fontsDir = Files.createTempDirectory("pdf-export-fonts");
        Files.copy(ARABIC_FONT.toPath(), fontsDir.resolve(ARABIC_FONT.getName()));
        Files.writeString(fontsDir.resolve("not-a-font.txt"), "not a font");
        Files.writeString(fontsDir.resolve("bogus.ttf"), "not a real ttf");
        Files.createDirectory(fontsDir.resolve("nested.ttf"));

        final var pdfFile = docsPath.resolve("pdf-export-fonts-dir.pdf").toFile();
        pdfFile.getParentFile().mkdirs();
        saveAsPdf(pdfFile, wmlPackage, fontsDir);
        assertTrue(pdfFile.isFile());
        assertTrue(pdfFile.length() > 0);
        try (final var is = Files.newInputStream(pdfFile.toPath())) {
            final var header = new byte[4];
            assertEquals(is.read(header), 4);
            assertEquals(new String(header), "%PDF");
        }
    }

    @Test
    public void testRegisterFontsDirectoryIgnoresNonFontFiles() throws Exception {
        checkFontsAvailable();
        final var fontsDir = Files.createTempDirectory("pdf-export-fonts");
        Files.copy(ARABIC_FONT.toPath(), fontsDir.resolve(ARABIC_FONT.getName()));
        Files.writeString(fontsDir.resolve("notes.txt"), "ignored");
        Files.writeString(fontsDir.resolve("fake.otf"), "ignored");
        // must not throw
        registerFonts(fontsDir);
        // null and non-directory arguments are ignored too
        registerFonts((Path) null);
        registerFonts(fontsDir.resolve("notes.txt"));
    }

    @Test
    public void testMissingFont() throws Exception {
        final var wmlPackage = loadWmlPackage();
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
