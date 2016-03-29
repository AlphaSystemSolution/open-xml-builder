/**
 *
 */
package com.alphasystem.openxml.builder.wml;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.TcPrInner.GridSpan;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.lang.Thread.currentThread;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.XmlUtils.unmarshal;
import static org.docx4j.openpackaging.parts.relationships.Namespaces.NS_WORD12;
import static org.docx4j.wml.STBorder.NONE;
import static org.docx4j.wml.STBrType.PAGE;
import static org.docx4j.wml.STFldCharType.BEGIN;
import static org.docx4j.wml.STFldCharType.END;

/**
 * @author sali
 */
public class WmlAdapter {

    public static final Text SINGLE_SPACE = getText(" ", "preserve");

    private static final Logger LOGGER = LoggerFactory.getLogger(WmlAdapter.class);

    private static final ClassLoader CLASS_LOADER = currentThread().getContextClassLoader();

    private static final AtomicInteger bookmarkCount = new AtomicInteger(0);

    static Styles loadStyles(Styles styles, String... paths) {
        if (isEmpty(paths)) {
            return styles;
        }
        for (String p : paths) {
            try {
                final List<URL> urls = readResources(p);
                if (urls.isEmpty()) {
                    return styles;
                }
                for (URL url : urls) {
                    System.out.println(format("Reading style {%s}", url));
                    try (InputStream ins = url.openStream()) {
                        final Styles otherStyles = (Styles) unmarshal(ins);
                        if (styles == null) {
                            styles = otherStyles;
                        } else {
                            styles.getStyle().addAll(otherStyles.getStyle());
                        }
                    } catch (JAXBException e) {
                        e.printStackTrace();
                    }
                }
            } catch (IOException e) {
                throw new UncheckedIOException(e.getMessage(), e);
            }
        } // end of for
        return styles;
    }

    static Numbering loadNumbering(Numbering numbering, String... paths) {
        if (isEmpty(paths)) {
            return null;
        }
        for (String p : paths) {
            try {
                final List<URL> urls = readResources(p);
                if (urls.isEmpty()) {
                    return numbering;
                }
                for (URL url : urls) {
                    System.out.println(format("Reading numbering {%s}", url));
                    try (InputStream ins = url.openStream()) {
                        final Numbering o = (Numbering) unmarshal(ins);
                        if (numbering == null) {
                            numbering = o;
                        } else {
                            numbering.getAbstractNum().addAll(o.getAbstractNum());
                            numbering.getNum().addAll(o.getNum());
                        }
                    } catch (JAXBException e) {
                        e.printStackTrace();
                    }
                }
            } catch (IOException e) {
                throw new UncheckedIOException(e.getMessage(), e);
            }
        }
        return numbering;
    }

    static List<URL> readResources(String pathName) throws IOException {
        LOGGER.info("Loading resource \"{}\".", pathName);
        List<URL> urls = new ArrayList<>();
        Path path = Paths.get(pathName);
        if (Files.exists(path)) {
            // path is a absolute file
            LOGGER.info("Path \"{}\" is an absolute file.", pathName);
            urls.add(path.toUri().toURL());
        } else {
            Enumeration<URL> resources;
            resources = CLASS_LOADER.getResources(pathName);
            if (resources == null || !resources.hasMoreElements()) {
                LOGGER.info("Resource \"{}\" does not exists.", pathName);
                pathName = format("META-INF/%s", pathName);
                resources = CLASS_LOADER.getResources(pathName);
            }
            if (resources == null || !resources.hasMoreElements()) {
                LOGGER.info("Resource \"{}\" does not exists.", pathName);
                return urls;
            }
            while (resources.hasMoreElements()) {
                urls.add(resources.nextElement());
            }
            Collections.reverse(urls);
        }
        return urls;
    }

    private static void addFieldBegin(P paragraph) {
        FldChar fldchar = getFldCharBuilder().withDirty(true).withFldCharType(BEGIN).getObject();
        R r = getRBuilder().addContent(getWrappedFldChar(fldchar)).getObject();
        paragraph.getContent().add(r);
    }

    private static void addFieldEnd(P paragraph) {
        FldChar fldchar = getFldCharBuilder().withFldCharType(END).getObject();
        R r = getRBuilder().addContent(getWrappedFldChar(fldchar)).getObject();
        paragraph.getContent().add(r);
    }

    private static void addTableOfContentField(P paragraph, String tocText) {
        final ObjectFactory objectFactory = Context.getWmlObjectFactory();
        Text txt = getText(tocText, "preserve");
        objectFactory.createRInstrText(txt);
        paragraph.getContent().add(getRBuilder().addContent(objectFactory.createRInstrText(txt)).getObject());
    }

    public static void addBookMark(PBuilder pBuilder, String id) {
        if (isBlank(id)) {
            return;
        }
        final long value = bookmarkCount.longValue();
        final CTBookmark bookmarkStart = getCTBookmarkBuilder().withId(value).withName(id).getObject();
        final JAXBElement<CTMarkupRange> bookmarkEnd = createCTMarkupRange(getCTBookmarkRangeBuilder().withId(value).getObject());
        pBuilder.getObject().getContent().add(0, bookmarkStart);
        pBuilder.addContent(bookmarkEnd);
    }

    public static void addBookMark(P p, String id) {
        if (isBlank(id)) {
            return;
        }
        String bookmarkId = String.valueOf(bookmarkCount.incrementAndGet());
        CTBookmark bookmarkStart = getCTBookmarkBuilder().withId(bookmarkId).withName(id).getObject();
        JAXBElement<CTMarkupRange> bookmarkEnd = createCTMarkupRange(getCTBookmarkRangeBuilder().withId(bookmarkId).getObject());
        p.getContent().add(0, bookmarkStart);
        p.getContent().add(bookmarkEnd);
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
        final PPrBuilder pPrBuilder = getPPrBuilder();
        PPrBase.NumPr numPr = applyNumbering ? getNumPr(listId, level) : null;
        return pPrBuilder.withNumPr(numPr).withPStyle(style).getObject();
    }

    public static CTLongHexNumber getCtLongHexNumber(String value) {
        return (value == null) ? null : getCTLongHexNumberBuilder().withVal(value).getObject();
    }

    /**
     * @param value
     * @return
     */
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

    /**
     * @param value
     * @return
     */
    public static GridSpan getGridSpan(long value) {
        return getTcPrInnerBuilder().getGridSpanBuilder().withVal(value).getObject();
    }

    /**
     * @param value
     * @return
     */
    public static HpsMeasure getHpsMeasure(long value) {
        return getHpsMeasureBuilder().withVal(value).getObject();
    }

    /**
     * @param value
     * @return
     */
    public static HpsMeasure getHpsMeasure(String value) {
        return getHpsMeasureBuilder().withVal(value).getObject();
    }

    /**
     * @param styleName
     * @return
     */
    public static PStyle getPStyle(String styleName) {
        return getPPrBuilder().getPStyleBuilder().withVal(styleName).getObject();
    }

    /**
     * @param styleName
     * @return
     */
    public static RStyle getRStyle(String styleName) {
        return getRStyleBuilder().withVal(styleName).getObject();
    }

    /**
     * @param value
     * @return
     */
    public static TblGridCol getTblGridCol(long value) {
        return getTblGridColBuilder().withW(value).getObject();
    }

    /**
     * @param value
     * @return
     */
    public static TblGridCol getTblGridCol(String value) {
        return getTblGridColBuilder().withW(value).getObject();
    }

    public static CTBorder getNilBorder() {
        return getBorder(NONE, 0L, 0L, "auto");
    }

    public static CTBorder getBorder(STBorder borderType, Long size, Long space, String color) {
        return getCTBorderBuilder().withVal(borderType).withSz(size).withSpace(space).withColor(color).getObject();
    }

    /**
     * @param value
     * @param space
     * @return
     */
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

    public static P getParagraphWithStyle(String styleName, String text) {
        String rsidRPr = nextId();
        R r = getRBuilder().withRsidRPr(rsidRPr).addContent(getText(text)).getObject();
        return getPBuilder(getEmptyPara(styleName)).withRsidRPr(rsidRPr).addContent(r).getObject();
    }

    public static P getHorizontalLine() {
        PPrBase.PBdr pBdr = WmlBuilderFactory.getPPrBaseBuilder().getPBdrBuilder().
                withTop(getBorder(STBorder.SINGLE, 6L, 1L, "auto")).getObject();
        PPr pPr = getPPrBuilder().withPBdr(pBdr).getObject();
        return getPBuilder().withPPr(pPr).getObject();
    }

    public static P getPageBreak() {
        String id = nextId();
        Br br = getBrBuilder().withType(PAGE).getObject();
        R r = getRBuilder().addContent(br).getObject();
        return getPBuilder().withRsidP(id).withRsidR(id).withRsidRDefault(id).addContent(r).getObject();
    }

    private static P addTableOfContentInternal(String tocText) {
        final PPr pPr = getPPrBuilder().withPStyle(getPStyle("TOC1")).getObject();
        P p = getPBuilder().withParaId(nextId()).withRsidR(nextId()).withRsidRDefault(nextId()).withRsidP(nextId())
                .withPPr(pPr).getObject();
        addFieldBegin(p);
        addTableOfContentField(p, tocText);
        addFieldEnd(p);
        return p;
    }

    private static P addTableOfContentTitle(String tocTitle) {
        final PPrBuilder pPrBuilder = getPPrBuilder();
        final PPrBase.NumPr numPr = pPrBuilder.getNumPrBuilder().withIlvl(0L).getObject();
        final PPr pPr = pPrBuilder.withPStyle("TOCHeading").withNumPr(numPr).getObject();
        final Text text = getText(tocTitle);
        final R r = WmlBuilderFactory.getRBuilder().addContent(text).getObject();
        return getPBuilder().withPPr(pPr).addContent(r).getObject();
    }

    public static List<P> addTableOfContent(String tocTitle, String tocText) {
        List<P> paras = new ArrayList<>();
        paras.add(addTableOfContentTitle(tocTitle));
        paras.add(addTableOfContentInternal(tocText));
        paras.add(getPageBreak());
        return paras;
    }

    public static List<P> addTableOfContent(String tocTitle) {
        return addTableOfContent(tocTitle, " TOC \\o \"1-3\" \\h \\z \\u \\h ");
    }

    public static List<P> addTableOfContent() {
        return addTableOfContent("Table of Contents");
    }

    public static void save(File file, WordprocessingMLPackage wordMLPackage)
            throws Docx4JException {
        wordMLPackage.save(getFile(file, "docx"));
    }

    public static void saveAsPdf(File file,
                                 WordprocessingMLPackage wordMLPackage, String fontName,
                                 boolean saveFO) throws Exception {
        File pdfFile = getFile(file, "pdf");

        // Set up font mapper (optional)
        Mapper fontMapper = new IdentityPlusMapper();
        wordMLPackage.setFontMapper(fontMapper);

        PhysicalFont font = PhysicalFonts.get(fontName);
        if (font != null) {
            fontMapper.put("Times New Roman", font);
            fontMapper.put("Arial", font);
        }

        FOSettings foSettings = Docx4J.createFOSettings();
        if (saveFO) {
            foSettings.setFoDumpFile(getFile(pdfFile, "fo"));
        }
        foSettings.setWmlPackage(wordMLPackage);

        OutputStream os = new java.io.FileOutputStream(pdfFile);

        // Don't care what type of exporter you use
        Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
    }
}
