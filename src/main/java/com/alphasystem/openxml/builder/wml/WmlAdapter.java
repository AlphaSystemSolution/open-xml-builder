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
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.TcPrInner.GridSpan;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.List;

import static com.alphasystem.openxml.builder.OpenXmlBuilder.OBJECT_FACTORY;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.lang.String.valueOf;
import static java.lang.Thread.currentThread;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.XmlUtils.unmarshal;
import static org.docx4j.openpackaging.parts.relationships.Namespaces.NS_WORD12;
import static org.docx4j.wml.STBrType.PAGE;

/**
 * @author sali
 */
public class WmlAdapter {

    public static final Text SINGLE_SPACE = getText(" ", "preserve");

    private static final ClassLoader CLASS_LOADER = currentThread().getContextClassLoader();

    static Styles loadStyles(Styles styles, String... paths) {
        if (isEmpty(paths)) {
            return null;
        }
        for (String p : paths) {
            try {
                final List<URL> urls = readResources(p);
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

    private static List<URL> readResources(String path) throws IOException {
        List<URL> urls = new ArrayList<>();
        Enumeration<URL> resources = CLASS_LOADER.getResources(format("META-INF/%s", path));
        if (resources != null) {
            while (resources.hasMoreElements()) {
                urls.add(resources.nextElement());
            }
        }
        Collections.reverse(urls);
        return urls;
    }

    public static JAXBElement<CTBookmark> createBodyBookmarkStart(
            CTBookmark bookmark) {
        return OBJECT_FACTORY.createBodyBookmarkStart(bookmark);
    }

    public static JAXBElement<CTMarkupRange> createCTMarkupRange(
            CTMarkupRange value) {
        return OBJECT_FACTORY.createBodyBookmarkEnd(value);
    }

    public static PPr getListParagraphProperties(long listId, long level) {
        final PPrBaseNumPrIlvlBuilder ilvlBuilder = getPPrBaseNumPrIlvlBuilder().withVal(valueOf(level));
        final PPrBaseNumPrNumIdBuilder numIdBuilder = getPPrBaseNumPrNumIdBuilder().withVal(valueOf(listId));
        final PPrBaseNumPrBuilder numPrBuilder = getPPrBaseNumPrBuilder().withIlvl(ilvlBuilder.getObject())
                .withNumId(numIdBuilder.getObject());
        return getPPrBuilder().withNumPr(numPrBuilder.getObject()).withPStyle(getPStyle("ListParagraph")).getObject();
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
    public static GridSpan getGridSpan(int value) {
        return getGridSpan(Integer.toString(value));
    }

    /**
     * @param value
     * @return
     */
    public static GridSpan getGridSpan(String value) {
        return getTcPrInnerGridSpanBuilder().withVal(value).getObject();
    }

    /**
     * @param value
     * @return
     */
    public static HpsMeasure getHpsMeasure(Integer value) {
        return getHpsMeasureBuilder().withVal(Integer.toString(value))
                .getObject();
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
        return getPPrBasePStyleBuilder().withVal(styleName).getObject();
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
    public static TblGridCol getTblGridCol(int value) {
        return getTblGridColBuilder().withW(Integer.toString(value))
                .getObject();
    }

    /**
     * @param value
     * @return
     */
    public static TblGridCol getTblGridCol(String value) {
        return getTblGridColBuilder().withW(value).getObject();
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
        return new JAXBElement(new QName(NS_WORD12, "fldChar"), FldChar.class,
                fldchar);
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
        styleName = isBlank(styleName) ? "Normal" : styleName;
        PStyle style = getPPrBasePStyleBuilder().withVal(styleName).getObject();
        if (style != null) {
            ppr = new PPrBuilder().withPStyle(style).getObject();
        }
        return getPBuilder().withPPr(ppr).withRsidP(id).withRsidR(id).withRsidRDefault(id).getObject();
    }

    public static P getPageBreak() {
        String id = nextId();
        Br br = getBrBuilder().withType(PAGE).getObject();
        R r = getRBuilder().addContent(br).getObject();
        return getPBuilder().withRsidP(id).withRsidR(id).withRsidRDefault(id).addContent(r).getObject();
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
