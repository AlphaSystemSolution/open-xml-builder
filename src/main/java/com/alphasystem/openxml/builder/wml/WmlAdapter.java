/**
 *
 */
package com.alphasystem.openxml.builder.wml;

import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.TcPrInner.GridSpan;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.namespace.QName;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.Enumeration;

import static com.alphasystem.openxml.builder.OpenXmlBuilder.OBJECT_FACTORY;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.lang.Thread.currentThread;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.openpackaging.parts.relationships.Namespaces.NS_WORD12;
import static org.docx4j.wml.STBrType.PAGE;

/**
 * @author sali
 */
public class WmlAdapter {

    public static final Text SINGLE_SPACE = getText(" ", "preserve");

    /**
     * @param wordDoc
     * @param styleFilePrefix
     */
    private static void addCustomStyle(WordprocessingMLPackage wordDoc, String styleFilePrefix) {
        StyleDefinitionsPart sdp = wordDoc.getMainDocumentPart()
                .getStyleDefinitionsPart();
        Styles styles = null;
        ClassLoader contextClassLoader = currentThread().getContextClassLoader();
        try {
            Enumeration<URL> resources = contextClassLoader
                    .getResources(String.format("META-INF/%s.xml", styleFilePrefix));
            if (resources != null) {
                while (resources.hasMoreElements()) {
                    URL url = resources.nextElement();
                    InputStream ins = url.openStream();
                    if (ins != null) {
                        try {
                            if(styles == null) {
                                styles = (Styles) XmlUtils.unmarshal(ins);
                            } else {
                                final Styles otherStyles = (Styles) XmlUtils.unmarshal(ins);
                                styles.getStyle().addAll(otherStyles.getStyle());
                            }
                        } catch (JAXBException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        if(styles != null) {
            sdp.setJaxbElement(styles);
        }
    }

    public static JAXBElement<CTBookmark> createBodyBookmarkStart(
            CTBookmark bookmark) {
        return OBJECT_FACTORY.createBodyBookmarkStart(bookmark);
    }

    public static JAXBElement<CTMarkupRange> createCTMarkupRange(
            CTMarkupRange value) {
        return OBJECT_FACTORY.createBodyBookmarkEnd(value);
    }

    /**
     * @return
     * @throws InvalidFormatException
     */
    public static WordprocessingMLPackage createNewDoc()
            throws InvalidFormatException {
        return createNewDoc("styles");
    }

    public static WordprocessingMLPackage createNewDoc(String styleFilePrefix)
            throws InvalidFormatException {
        WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage
                .createPackage();
        addCustomStyle(wordprocessingMLPackage, styleFilePrefix);
        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ndp.setJaxbElement(NumberingHelper.createNumbering());
        wordprocessingMLPackage.getMainDocumentPart().addTargetPart(ndp);
        return wordprocessingMLPackage;
    }

    /**
     * @param customStyles
     * @return
     * @throws InvalidFormatException
     */
    public static WordprocessingMLPackage createNewDoc(Styles customStyles)
            throws InvalidFormatException {
        WordprocessingMLPackage wordDoc = WordprocessingMLPackage
                .createPackage();
        if (customStyles != null) {
            StyleDefinitionsPart sdp = wordDoc.getMainDocumentPart()
                    .getStyleDefinitionsPart();
            sdp.setJaxbElement(customStyles);
        }
        return wordDoc;
    }

    public static CTLongHexNumber getCtLongHexNumber(String value) {
        return getCTLongHexNumberBuilder().withVal(value).getObject();
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
