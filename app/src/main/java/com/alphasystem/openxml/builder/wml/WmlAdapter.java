/**
 *
 */
package com.alphasystem.openxml.builder.wml;


import com.alphasystem.docx4j.builder.wml.PBuilder;
import com.alphasystem.docx4j.builder.wml.PPrBaseBuilder;
import com.alphasystem.docx4j.builder.wml.WmlBuilderFactory;
import jakarta.xml.bind.JAXBElement;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.relationships.ObjectFactory;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.docx4j.wml.TcPrInner.GridSpan;

import javax.xml.namespace.QName;
import java.io.File;
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
        final long id = bookmarkCount.longValue();
        final CTBookmark bookmarkStart = getCTBookmarkBuilder().withId(id).withName(name).getObject();
        final JAXBElement<CTMarkupRange> bookmarkEnd = createCTMarkupRange(getCTBookmarkRangeBuilder().withId(id).getObject());
        pBuilder.getObject().getContent().add(0, bookmarkStart);
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

    public static P getPageBreak() {
        String id = nextId();
        Br br = getBrBuilder().withType(PAGE).getObject();
        R r = getRBuilder().addContent(br).getObject();
        return getPBuilder().withRsidP(id).withRsidR(id).withRsidRDefault(id).addContent(r).getObject();
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
}
