package com.alphasystem.openxml.builder.wml;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.wml.STFldCharType.*;
import static org.docx4j.wml.STTabJc.LEFT;
import static org.docx4j.wml.STTabJc.RIGHT;
import static org.docx4j.wml.STTabTlc.DOT;
import static org.docx4j.wml.STTheme.*;

/**
 * @author sali
 */
public final class TocGenerator {

    private static final ObjectFactory wmlObjectFactory = Context.getWmlObjectFactory();
    private static final String DEFAULT_TOC_HEADING_STYE = "TOCHeading";
    private static final String DEFAULT_TOC_STYLE = "TOC1";
    private static final String DEFAULT_TOC_HEADING = "Contents";
    private static final String INSTRUCTION = " TOC \\o \"1-%s\" \\h \\z \\u ";

    private int index;

    private String instruction;

    private String tocStyle;

    private String tocHeadingStyle;

    private String tocHeading;

    private MainDocumentPart mainDocumentPart;

    public TocGenerator() {
        index(0);
        level(3);
        tocStyle(null);
        tocHeading(null);
        tocHeadingStyle(null);
    }

    public TocGenerator index(int index) {
        this.index = Math.max(index, 0);
        return this;
    }

    public TocGenerator instruction(String instruction) {
        this.instruction = isBlank(instruction) ? INSTRUCTION : instruction;
        return this;
    }

    public TocGenerator level(int level) {
        level = ((level <= 0) || (level > 5)) ? 3 : level;
        return instruction(format(INSTRUCTION, level));
    }

    public TocGenerator tocStyle(String tocStyle) {
        this.tocStyle = isBlank(tocStyle) ? DEFAULT_TOC_STYLE : tocStyle;
        return this;
    }

    public TocGenerator tocHeadingStyle(String tocHeadingStyle) {
        this.tocHeadingStyle = isBlank(tocHeadingStyle) ? DEFAULT_TOC_HEADING_STYE : tocHeadingStyle;
        return this;
    }

    public TocGenerator tocHeading(String tocHeading) {
        this.tocHeading = isBlank(tocHeading) ? DEFAULT_TOC_HEADING : tocHeading;
        return this;
    }

    public TocGenerator mainDocumentPart(MainDocumentPart mainDocumentPart) {
        this.mainDocumentPart = mainDocumentPart;
        return this;
    }

    public void generateToc() {
        SdtPr sdtPr = createSdtPr();
        CTSdtEndPr ctSdtEndPr = createCtSdtEndPr();
        SdtContentBlock sdtContentBlock = createSdtContentBlock();
        SdtBlock sdtBlock = getSdtBlockBuilder().withSdtPr(sdtPr).withSdtEndPr(ctSdtEndPr).withSdtContent(sdtContentBlock).getObject();
        mainDocumentPart.getContent().add(index, sdtBlock);
        try {
            WmlAdapter.updateSettings(mainDocumentPart);
        } catch (InvalidFormatException e) {
            // ignore
            e.printStackTrace();
        }
        mainDocumentPart.getContent().add(index + 1, WmlAdapter.getPageBreak());
    }

    private SdtPr createSdtPr() {
        CTSdtDocPartBuilder ctSdtDocPartBuilder = WmlBuilderFactory.getCTSdtDocPartBuilder();
        CTSdtDocPartBuilder.DocPartGalleryBuilder docPartGalleryBuilder = ctSdtDocPartBuilder.getDocPartGalleryBuilder()
                .withVal("Table of Contents");
        ctSdtDocPartBuilder.withDocPartUnique(true).withDocPartGallery(docPartGalleryBuilder.getObject());

        JAXBElement<CTSdtDocPart> sdtdocpartWrapped = wmlObjectFactory.createSdtPrDocPartObj(ctSdtDocPartBuilder.getObject());

        SdtPrBuilder sdtPrBuilder = WmlBuilderFactory.getSdtPrBuilder().addRPrOrAliasOrLock(sdtdocpartWrapped);
        SdtPr sdtPr = sdtPrBuilder.getObject();
        sdtPr.setId();
        return sdtPr;
    }

    private CTSdtEndPr createCtSdtEndPr() {
        RFontsBuilder rFontsBuilder = getRFontsBuilder().withAsciiTheme(MINOR_H_ANSI).withEastAsiaTheme(MINOR_H_ANSI)
                .withHAnsiTheme(MINOR_H_ANSI).withCstheme(MINOR_BIDI);
        RPrBuilder rPrBuilder = getRPrBuilder().withRFonts(rFontsBuilder.getObject()).withB(true).withBCs(true)
                .withNoProof(true).withSz(22L).withSzCs(22L);
        return getCTSdtEndPrBuilder().addRPr(rPrBuilder.getObject()).getObject();
    }

    private SdtContentBlock createSdtContentBlock() {
        P headingP = WmlAdapter.getParagraphWithStyle(tocHeadingStyle, tocHeading);
        WmlAdapter.addBookMark(headingP, tocHeading.replaceAll(" ", "_").toLowerCase());
        P contentP = getSdtContentP();
        P fldEndP = getFieldEndPara();
        return getSdtContentBlockBuilder().addContent(headingP, contentP, fldEndP).getObject();
    }

    private P getSdtContentP() {
        CTTabStop leftTab = getCTTabStopBuilder().withVal(LEFT).withPos(440L).getObject();
        CTTabStop rightTab = getCTTabStopBuilder().withVal(RIGHT).withPos(9350L).withLeader(DOT).getObject();
        Tabs tabs = getTabsBuilder().addTab(leftTab, rightTab).getObject();
        PPrBuilder pPrBuilder = getPPrBuilder().withPStyle(tocStyle).withTabs(tabs);
        return getPBuilder().withPPr(pPrBuilder.getObject()).addContent(getFieldBeginRun(), getInstructionRun(instruction),
                getFieldSeparateRun()).getObject();
    }

    private static R getFieldBeginRun() {
        FldChar fldchar = getFldCharBuilder().withDirty(true).withFldCharType(BEGIN).getObject();
        return getRBuilder().addContent(WmlAdapter.getWrappedFldChar(fldchar)).getObject();
    }

    private static R getInstructionRun(String instruction) {
        Text txt = WmlAdapter.getText(instruction, "preserve");
        wmlObjectFactory.createRInstrText(txt);
        return getRBuilder().addContent(wmlObjectFactory.createRInstrText(txt)).getObject();
    }


    private static R getFieldSeparateRun() {
        FldChar fldchar = getFldCharBuilder().withDirty(true).withFldCharType(SEPARATE).getObject();
        return getRBuilder().addContent(WmlAdapter.getWrappedFldChar(fldchar)).getObject();
    }

    private static P getFieldEndPara() {
        FldChar fldchar = getFldCharBuilder().withFldCharType(END).getObject();
        R r = getRBuilder().addContent(WmlAdapter.getWrappedFldChar(fldchar)).getObject();
        return getPBuilder().addContent(r).getObject();
    }

}
