/**
 *
 */
package com.alphasystem.docx4j.builder.wml;


import jakarta.xml.bind.JAXBElement;
import org.docx4j.Docx4J;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.ObjectFactory;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.TcPrInner.GridSpan;

import javax.xml.namespace.QName;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Locale;
import java.util.List;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

import static com.alphasystem.commons.util.IdGenerator.nextId;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.openpackaging.parts.relationships.Namespaces.HYPERLINK;
import static org.docx4j.openpackaging.parts.relationships.Namespaces.NS_WORD12;
import static org.docx4j.wml.STBorder.NIL;
import static org.docx4j.wml.STBorder.SINGLE;
import static org.docx4j.wml.STBrType.PAGE;

/**
 * @author sali
 */
public class WmlAdapter {

    public static final Text SINGLE_SPACE = getText(" ", "preserve");

    private static final AtomicInteger bookmarkCount = new AtomicInteger(0);

    public static String addExternalHyperlinkRelationship(String url, MainDocumentPart mainDocumentPart) {
        final org.docx4j.relationships.ObjectFactory objectFactory = new ObjectFactory();
        final Relationship relationship = objectFactory.createRelationship();
        relationship.setType(HYPERLINK);
        relationship.setTarget(url);
        relationship.setTargetMode("External");
        mainDocumentPart.getRelationshipsPart().addRelationship(relationship);
        return relationship.getId();
    }

    public static P.Hyperlink addHyperlink(String href, String linkText) {
        final RPr rPr = getRPrBuilder().withRStyle("Hyperlink").getObject();
        final R run = getRBuilder().withRPr(rPr).addContent(getText(linkText)).getObject();
        return getPHyperlinkBuilder().withHistory(true).withAnchor(href).addContent(run).getObject();
    }

    public static P.Hyperlink addExternalHyperlink(String href,
                                                   String linkText,
                                                   MainDocumentPart mainDocumentPart) {
        final RPr rPr = getRPrBuilder().withRStyle("Hyperlink").getObject();
        final R run = getRBuilder().withRPr(rPr).addContent(getText(linkText)).getObject();
        return getPHyperlinkBuilder().withHistory(true).withId(addExternalHyperlinkRelationship(href, mainDocumentPart))
                .addContent(run).getObject();
    }

    public static void addBookMark(PBuilder pBuilder, String name) {
        if (isBlank(name)) {
            return;
        }
        final long id = bookmarkCount.getAndIncrement();
        final CTBookmark bookmarkStart = getCTBookmarkBuilder().withId(id).withName(name).getObject();
        final JAXBElement<CTMarkupRange> bookmarkEnd = createCTMarkupRange(getCTBookmarkRangeBuilder().withId(id).getObject());
        pBuilder.getObject().getContent().addFirst(bookmarkStart);
        pBuilder.addContent(bookmarkEnd);
    }

    public static void addBookMark(P p, String name) {
        if (isBlank(name)) {
            return;
        }
        addBookMark(getPBuilder(p), name);
    }

    public static JAXBElement<CTBookmark> createBodyBookmarkStart(
            CTBookmark bookmark) {
        return OBJECT_FACTORY.createBodyBookmarkStart(bookmark);
    }

    public static JAXBElement<CTMarkupRange> createCTMarkupRange(
            CTMarkupRange value) {
        return OBJECT_FACTORY.createBodyBookmarkEnd(value);
    }

    public static PPrBase.NumPr getNumPr(Long numId, Long level) {
        return new PPrBaseBuilder().getNumPrBuilder().withNumId(numId).withIlvl(level).getObject();
    }

    public static PPr getListParagraphProperties(long listId, long level, boolean applyNumbering) {
        return getListParagraphProperties(listId, level, "ListParagraph", applyNumbering);
    }

    public static PPr getListParagraphProperties(long listId, long level, String style, boolean applyNumbering) {
        style = isBlank(style) ? "ListParagraph" : style;
        final var pPrBuilder = getPPrBuilder();
        PPrBase.NumPr numPr = applyNumbering ? getNumPr(listId, level) : null;
        return pPrBuilder.withNumPr(numPr).withPStyle(style).getObject();
    }

    public static PPr getListParagraphProperties(ListItem<?> listItem, long level, boolean applyNumbering) {
        return getListParagraphProperties(listItem, level, null, applyNumbering);
    }

    public static PPr getListParagraphProperties(ListItem<?> listItem, long level, String style, boolean applyNumbering) {
        final var sanitizeStyle = isBlank(style) ? "ListParagraph" : style;
        final var numPr = applyNumbering ? getNumPr(listItem.getNumberId(), level) : null;
        return getPPrBuilder().withPStyle(sanitizeStyle).withNumPr(numPr).getObject();
    }

    public static Color getColor(String value) {
        return getColorBuilder().withVal(value).getObject();
    }

    public static File getFile(File file, String extension) {
        String fileName = file.getName();
        if (fileName.endsWith(extension)) {
            return file;
        }
        File folder = file.getParentFile();
        int index = fileName.lastIndexOf('.');
        if (index > -1) {
            fileName = fileName.substring(0, index);
        }
        return new File(folder, format("%s.%s", fileName, extension));
    }

    public static GridSpan getGridSpan(long value) {
        return getTcPrInnerBuilder().getGridSpanBuilder().withVal(value).getObject();
    }

    public static CTBorder getDefaultBorder() {
        return getBorder(SINGLE, 0L, 0L, "auto");
    }

    public static CTBorder getNilBorder() {
        return getBorder(NIL, 0L, 0L, "auto");
    }

    public static CTBorder getBorder(STBorder borderType, Long size, Long space, String color) {
        return getCTBorderBuilder().withVal(borderType).withSz(size).withSpace(space).withColor(color).getObject();
    }

    public static TblBorders getNilTblBorders() {
        return getTblBordersBuilder()
                .withTop(getNilBorder())
                .withBottom(getNilBorder())
                .withLeft(getNilBorder())
                .withRight(getNilBorder())
                .withInsideH(getNilBorder())
                .withInsideV(getNilBorder())
                .getObject();
    }

    public static TcPrInner.TcBorders getNilBorders() {
        return getTcPrBuilder().getTcBordersBuilder().withTop(getNilBorder()).withBottom(getNilBorder())
                .withLeft(getNilBorder()).withRight(getNilBorder()).withInsideH(getNilBorder()).withInsideV(getNilBorder())
                .getObject();
    }

    public static TcPrInner.TcBorders getDefaultBorders() {
        return getTcPrBuilder().getTcBordersBuilder().withTop(getDefaultBorder()).withBottom(getDefaultBorder())
                .withLeft(getDefaultBorder()).withRight(getDefaultBorder()).withInsideH(getDefaultBorder())
                .withInsideV(getDefaultBorder()).getObject();
    }

    public static Text getText(String value, String space) {
        return getTextBuilder().withValue(value).withSpace(space).getObject();
    }

    public static Text getText(String value) {
        String space = !value.equals(value.trim()) ? "preserve" : null;
        return getText(value, space);
    }

    @SuppressWarnings({"unchecked", "rawtypes"})
    public static JAXBElement getWrappedFldChar(FldChar fldchar) {
        return new JAXBElement(new QName(NS_WORD12, "fldChar"), FldChar.class, fldchar);
    }

    public static P getEmptyPara() {
        return getEmptyPara(null);
    }

    public static P getEmptyParaNoSpacing() {
        return getEmptyPara("NoSpacing");
    }

    public static P getEmptyPara(String styleName) {
        String id = nextId();

        PPr ppr = null;
        if (styleName != null) {
            ppr = getPPrBuilder().withPStyle(styleName).getObject();
        }

        return getPBuilder().withPPr(ppr).withRsidP(id).withRsidR(id).withRsidRDefault(id).getObject();
    }

    public static P getParagraph(String text) {
        return getParagraphWithStyle(null, text);
    }

    public static P getParagraphWithStyle(String styleName, String text) {
        String rsidRPr = nextId();
        R r = getRBuilder().withRsidRPr(rsidRPr).addContent(getText(text)).getObject();
        return getPBuilder(getEmptyPara(styleName)).withRsidRPr(rsidRPr).addContent(r).getObject();
    }

    public static P getHorizontalLine() {
        PPrBase.PBdr pBdr = WmlBuilderFactory.getPPrBaseBuilder().getPBdrBuilder().
                withTop(getBorder(SINGLE, 6L, 1L, "auto")).getObject();
        PPr pPr = getPPrBuilder().withPBdr(pBdr).getObject();
        return getPBuilder().withPPr(pPr).getObject();
    }

    public static Br getPageBreak() {
        return getBrBuilder().withType(PAGE).getObject();
    }

    public static R getPageBreakRun() {
        return getRBuilder().addContent(getPageBreak()).getObject();
    }

    public static P getPageBreakPara() {
        final var id = nextId();
        return getPBuilder().withRsidP(id).withRsidR(id).withRsidRDefault(id).addContent(getPageBreakRun()).getObject();
    }

    public static P insertImage(WordprocessingMLPackage wordMLPackage, String filenameHint, String altText,
                                long id1, int id2, URL url) throws Exception {
        final var imagePart = BinaryPartAbstractImage.createLinkedImagePart(wordMLPackage, url);
        final var imageInline = imagePart.createImageInline(filenameHint, altText, id1, id2, false);
        final var drawing = getDrawingBuilder().addAnchorOrInline(imageInline).getObject();
        return getPBuilder().addContent(getRBuilder().addContent(drawing).getObject()).getObject();
    }

    static void updateSettings(MainDocumentPart mainDocumentPart) throws InvalidFormatException {
        //Adding Print View and Setting Update Field to true
        DocumentSettingsPart dsp = mainDocumentPart.getDocumentSettingsPart(true);
        CTSettings ct = new CTSettings();
        CTView ctView = Context.getWmlObjectFactory().createCTView();
        ctView.setVal(STView.PRINT);
        ct.setView(ctView);
        ct.setUpdateFields(WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE);
        dsp.setJaxbElement(ct);
        mainDocumentPart.addTargetPart(dsp);
    }

    public static void save(File file, WordprocessingMLPackage wordMLPackage)
            throws Docx4JException {
        wordMLPackage.save(getFile(file, "docx"));
    }

    /**
     * Registers the given font files so they can be resolved and embedded
     * during PDF conversion. Fonts are registered under the font's internal
     * name; use {@link #registerFont(String, File)} when the name used in
     * the document's styles differs from the font's internal name.
     *
     * @param fontFiles .ttf/.otf font files
     * @throws Docx4JException if a font file does not exist or cannot be loaded
     */
    public static void registerFonts(File... fontFiles) throws Docx4JException {
        if (fontFiles == null) {
            return;
        }
        for (File fontFile : fontFiles) {
            registerFont(null, fontFile);
        }
    }

    /**
     * Registers a font file, optionally under a specific font name as used
     * in the document's {@code w:rFonts} attributes.
     *
     * @param fontNameInDocx the font name used in the document, or {@code null}
     *                       to use the font's internal name
     * @param fontFile       .ttf/.otf font file
     * @throws Docx4JException if the font file does not exist or cannot be loaded
     */
    public static void registerFont(String fontNameInDocx, File fontFile) throws Docx4JException {
        if (fontFile == null || !fontFile.isFile()) {
            throw new Docx4JException(format("Font file does not exist: %s", fontFile));
        }
        PhysicalFonts.addPhysicalFonts(fontNameInDocx, fontFile.toURI());
    }

    /**
     * Registers all font files ({@code .ttf}, {@code .otf}, {@code .ttc})
     * found directly in the given directory. Files that are not font files,
     * and files that cannot be loaded as fonts, are silently ignored.
     *
     * @param fontsDirectory directory containing font files; ignored if
     *                       {@code null} or not a directory
     * @throws Docx4JException if the directory cannot be listed
     */
    public static void registerFonts(Path fontsDirectory) throws Docx4JException {
        if (fontsDirectory == null || !Files.isDirectory(fontsDirectory)) {
            return;
        }
        try (final var stream = Files.list(fontsDirectory)) {
            stream.filter(Files::isRegularFile)
                    .filter(WmlAdapter::isFontFile)
                    .forEach(path -> {
                        try {
                            registerFont(null, path.toFile());
                        } catch (Docx4JException | RuntimeException ignored) {
                            // not a loadable font file; silently ignore
                        }
                    });
        } catch (IOException e) {
            throw new Docx4JException(format("Failed to list fonts directory: %s", fontsDirectory), e);
        }
    }

    private static boolean isFontFile(Path path) {
        final var name = path.getFileName().toString().toLowerCase(Locale.ROOT);
        return name.endsWith(".ttf") || name.endsWith(".otf") || name.endsWith(".ttc");
    }

    /**
     * Converts the given package to PDF and saves it next to the given file
     * with a {@code .pdf} extension.
     *
     * @param file       target file (extension is replaced with {@code .pdf})
     * @param wmlPackage the package to convert
     * @param fontFiles  optional custom font files to register before converting
     * @throws Docx4JException if conversion fails or a used font cannot be resolved
     */
    public static void saveAsPdf(File file, WordprocessingMLPackage wmlPackage, File... fontFiles)
            throws Docx4JException {
        final var pdfFile = getFile(file, "pdf");
        try (final var os = new FileOutputStream(pdfFile)) {
            toPdf(wmlPackage, os, fontFiles);
        } catch (IOException e) {
            pdfFile.delete();
            throw new Docx4JException(e.getMessage(), e);
        } catch (Docx4JException | RuntimeException e) {
            pdfFile.delete();
            throw e;
        }
    }

    /**
     * Loads the given docx file and converts it to PDF, saving it next to
     * the given file with a {@code .pdf} extension.
     *
     * @param docxFile  the docx file to load
     * @param fontFiles optional custom font files to register before converting
     * @throws Docx4JException if loading/conversion fails or a used font cannot be resolved
     */
    public static void saveAsPdf(File docxFile, File... fontFiles) throws Docx4JException {
        saveAsPdf(docxFile, WordprocessingMLPackage.load(docxFile), fontFiles);
    }

    /**
     * Converts the given package to PDF and saves it next to the given file
     * with a {@code .pdf} extension, registering all font files found in the
     * given directory first. Non-font files in the directory are silently
     * ignored.
     *
     * @param file           target file (extension is replaced with {@code .pdf})
     * @param wmlPackage     the package to convert
     * @param fontsDirectory directory containing .ttf/.otf/.ttc font files
     * @throws Docx4JException if conversion fails or a used font cannot be resolved
     */
    public static void saveAsPdf(File file, WordprocessingMLPackage wmlPackage, Path fontsDirectory)
            throws Docx4JException {
        final var pdfFile = getFile(file, "pdf");
        try (final var os = new FileOutputStream(pdfFile)) {
            toPdf(wmlPackage, os, fontsDirectory);
        } catch (IOException e) {
            pdfFile.delete();
            throw new Docx4JException(e.getMessage(), e);
        } catch (Docx4JException | RuntimeException e) {
            pdfFile.delete();
            throw e;
        }
    }

    /**
     * Loads the given docx file and converts it to PDF, saving it next to
     * the given file with a {@code .pdf} extension, registering all font
     * files found in the given directory first. Non-font files in the
     * directory are silently ignored.
     *
     * @param docxFile       the docx file to load
     * @param fontsDirectory directory containing .ttf/.otf/.ttc font files
     * @throws Docx4JException if loading/conversion fails or a used font cannot be resolved
     */
    public static void saveAsPdf(File docxFile, Path fontsDirectory) throws Docx4JException {
        saveAsPdf(docxFile, WordprocessingMLPackage.load(docxFile), fontsDirectory);
    }

    /**
     * Converts the given package to PDF, writing to the given stream.
     * <p>
     * Fonts are resolved from the given font files first and from
     * OS-installed fonts (OS font directories are always scanned). If a font
     * used by the document cannot be resolved from either source, a
     * {@link Docx4JException} is thrown naming the missing fonts.
     * <p>
     * Conversion uses the XSL-FO/Apache FOP pipeline via
     * {@link Docx4J#toPDF}.
     *
     * @param wmlPackage the package to convert
     * @param out        the stream to write the PDF to
     * @param fontFiles  optional custom font files to register before converting
     * @throws Docx4JException if conversion fails or a used font cannot be resolved
     */
    public static void toPdf(WordprocessingMLPackage wmlPackage, OutputStream out, File... fontFiles)
            throws Docx4JException {
        registerFonts(fontFiles);
        convertToPdf(wmlPackage, out);
    }

    /**
     * Converts the given package to PDF, writing to the given stream, after
     * registering all font files found in the given directory. Non-font
     * files in the directory are silently ignored.
     *
     * @param wmlPackage     the package to convert
     * @param out            the stream to write the PDF to
     * @param fontsDirectory directory containing .ttf/.otf/.ttc font files
     * @throws Docx4JException if conversion fails or a used font cannot be resolved
     * @see #toPdf(WordprocessingMLPackage, OutputStream, File...)
     */
    public static void toPdf(WordprocessingMLPackage wmlPackage, OutputStream out, Path fontsDirectory)
            throws Docx4JException {
        registerFonts(fontsDirectory);
        convertToPdf(wmlPackage, out);
    }

    private static void convertToPdf(WordprocessingMLPackage wmlPackage, OutputStream out)
            throws Docx4JException {
        validateFonts(wmlPackage);
        Docx4J.toPDF(wmlPackage, out, Docx4J.FLAG_EXPORT_PREFER_XSL);
    }

    private static void validateFonts(WordprocessingMLPackage wmlPackage) throws Docx4JException {
        final Set<String> fontsInUse = wmlPackage.getMainDocumentPart().fontsInUse();
        if (fontsInUse == null || fontsInUse.isEmpty()) {
            return;
        }
        // fonts handled internally by docx4j substitutes
        final Set<String> builtin = Set.of("Symbol", "Webdings", "Wingdings", "Wingdings 2", "Wingdings 3");
        final var fontMapper = wmlPackage.getFontMapper();
        final List<String> missing = new ArrayList<>();
        for (String fontName : fontsInUse) {
            if (builtin.contains(fontName)) {
                continue;
            }
            // fontMapper.get() triggers OS font discovery but resolves any
            // name to a substitute, so it never returns null; PhysicalFonts.get()
            // only returns fonts that actually exist
            fontMapper.get(fontName);
            if (PhysicalFonts.get(fontName) == null) {
                missing.add(fontName);
            }
        }
        if (!missing.isEmpty()) {
            throw new Docx4JException(format(
                    "Fonts not found in provided font files or OS font directories: %s. " +
                            "Supply the .ttf/.otf font file(s) via the fontFiles parameter.", missing));
        }
    }
}
