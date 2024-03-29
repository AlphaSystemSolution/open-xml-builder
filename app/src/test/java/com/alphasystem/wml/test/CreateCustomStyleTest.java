package com.alphasystem.wml.test;

import com.alphasystem.docx4j.builder.wml.*;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.Spacing;
import org.testng.annotations.Test;

import static com.alphasystem.commons.util.IdGenerator.nextId;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getBorder;
import static com.alphasystem.docx4j.builder.wml.WmlAdapter.getNilBorder;
import static org.docx4j.XmlUtils.marshaltoString;
import static org.docx4j.wml.STBorder.SINGLE;
import static org.docx4j.wml.STLineSpacingRule.AUTO;
import static org.docx4j.wml.STShd.CLEAR;
import static org.docx4j.wml.STTblStyleOverrideType.*;
import static org.docx4j.wml.STThemeColor.TEXT_1;
import static org.docx4j.wml.STVerticalJc.CENTER;

/**
 * @author sali
 */
public class CreateCustomStyleTest {

    @Test
    public void createAdmonitionTableStyle() {
        final var ctTblPrBaseBuilder = new CTTblPrBaseBuilder();
        TblBorders tblBorders = new TblBordersBuilder().withTop(getNilBorder()).withLeft(getNilBorder())
                .withBottom(getNilBorder()).withRight(getNilBorder()).withInsideH(getNilBorder()).withInsideV(getNilBorder())
                .getObject();
        CTTblPrBase tblPr = ctTblPrBaseBuilder.withTblBorders(tblBorders).getObject();

        CTTblStylePr[] ctTblStylePrs = new CTTblStylePr[2];

        // first col, "right" border and Vertical Align "Center"
        TcPrBuilder tcPrBuilder = new TcPrBuilder();
        final CTBorder border = getBorder(SINGLE, 2L, 2L, "2C2C2C");
        TcPrInner.TcBorders tcBorders = tcPrBuilder.getTcBordersBuilder().withRight(border).getObject();
        CTShd shade = getCTShdBuilder().withVal(CLEAR).withColor("auto").withFill("F8F8F7").getObject();
        TcPr tcPr = tcPrBuilder.withTcBorders(tcBorders).withVAlign(CENTER).withShd(shade).getObject();
        RPr rPr = new RPrBuilder().withRStyle("Strong").getObject();
        ctTblStylePrs[0] = new CTTblStylePrBuilder().withType(FIRST_COL).withRPr(rPr)
                .withPPr(getPPrBuilder().withJc(JcEnumeration.CENTER).getObject()).withTcPr(tcPr).getObject();

        // second column, "left" border
        tcPrBuilder = new TcPrBuilder();
        tcBorders = tcPrBuilder.getTcBordersBuilder().withLeft(border).getObject();
        tcPr = tcPrBuilder.withTcBorders(tcBorders).withShd(shade).getObject();
        ctTblStylePrs[1] = new CTTblStylePrBuilder().withType(LAST_COL).withTcPr(tcPr).getObject();

        final var pPrBuilder = new PPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        final var styleBuilder = getStyleBuilder().withType("table").withCustomStyle(true).withStyleId("AdmonitionTable")
                .withName("Admonition Table").withBasedOn("TableNormal").withUiPriority(47L).withRsid(nextId())
                .withPPr(pPr).withTblPr(tblPr).addTblStylePr(ctTblStylePrs);

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }

    @Test
    public void createHorizontalListTableStyle() {
        final PPrBuilder pPrBuilder = new PPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        final CTTblPrBaseBuilder ctTblPrBaseBuilder = new CTTblPrBaseBuilder();
        TblBorders tblBorders = new TblBordersBuilder().withTop(getNilBorder()).withLeft(getNilBorder())
                .withBottom(getNilBorder()).withRight(getNilBorder()).withInsideH(getNilBorder()).withInsideV(getNilBorder())
                .getObject();
        CTTblPrBase tblPr = ctTblPrBaseBuilder.withTblBorders(tblBorders).getObject();

        CTTblStylePr[] ctTblStylePrs = new CTTblStylePr[1];

        // first column
        RPr rPr = new RPrBuilder().withRStyle("Strong").getObject();
        ctTblStylePrs[0] = new CTTblStylePrBuilder().withType(FIRST_COL).withRPr(rPr).getObject();

        StyleBuilder styleBuilder = getStyleBuilder().withType("table").withCustomStyle(true).withStyleId("HorizontalList")
                .withName("Horizontal List").withBasedOn("TableNormal").withUiPriority(47L).withRsid(nextId())
                .withPPr(pPr).withTblPr(tblPr).addTblStylePr(ctTblStylePrs);

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }

    @Test
    public void createTableStyleWithHeader() {
        PPrBuilder pPrBuilder = getPPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        StyleBuilder styleBuilder = getStyleBuilder().withType("table").withCustomStyle(true).withStyleId("TableGrid1")
                .withName("Table Grid With Header").withBasedOn("TableNormal").withUiPriority(48L).withRsid(nextId())
                .withPPr(pPr).withTblPr(createTableProperties()).addTblStylePr(createTableStyle(FIRST_ROW));

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }

    @Test
    public void createTableStyleWithFooter() {
        PPrBuilder pPrBuilder = getPPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        StyleBuilder styleBuilder = getStyleBuilder().withType("table").withCustomStyle(true).withStyleId("TableGrid2")
                .withName("Table Grid With Footer").withBasedOn("TableNormal").withUiPriority(48L).withRsid(nextId())
                .withPPr(pPr).withTblPr(createTableProperties()).addTblStylePr(createTableStyle(LAST_ROW));

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }

    @Test
    public void createTableStyleWithHeaderAndFooter() {
        PPrBuilder pPrBuilder = getPPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        StyleBuilder styleBuilder = getStyleBuilder().withType("table").withCustomStyle(true).withStyleId("TableGrid3")
                .withName("Table Grid With Header And Footer").withBasedOn("TableNormal").withUiPriority(49L).withRsid(nextId())
                .withPPr(pPr).withTblPr(createTableProperties()).addTblStylePr(createTableStyle(FIRST_ROW),
                        createTableStyle(LAST_ROW));

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }

    private CTTblStylePr createTableStyle(STTblStyleOverrideType type) {
        RPr rPr = getRPrBuilder().withRStyle("Strong").getObject();
        PPr pPr = getPPrBuilder().withJc(JcEnumeration.CENTER).getObject();
        TcPr tcPr = getTcPrBuilder().withVAlign(CENTER).getObject();
        return getCTTblStylePrBuilder().withType(type).withRPr(rPr).withPPr(pPr).withTcPr(tcPr).getObject();
    }

    private CTTblPrBase createTableProperties() {
        CTBorder top = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        CTBorder left = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        CTBorder bottom = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        CTBorder right = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        CTBorder insideH = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        CTBorder insideV = getCTBorderBuilder().withVal(SINGLE).withSz(4L).withSpace(0L).withColor("2C2C2C")
                .withThemeColor(TEXT_1).getObject();
        final TblBorders tblBorders = WmlBuilderFactory.getTblBordersBuilder().withTop(top).withLeft(left)
                .withBottom(bottom).withRight(right).withInsideH(insideH).withInsideV(insideV).getObject();
        return getCTTblPrBaseBuilder().withTblBorders(tblBorders).getObject();
    }
}
