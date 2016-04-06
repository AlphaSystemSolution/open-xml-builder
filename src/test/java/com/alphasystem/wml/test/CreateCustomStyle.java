package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.*;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.Spacing;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static org.docx4j.XmlUtils.marshaltoString;
import static org.docx4j.wml.STBorder.SINGLE;
import static org.docx4j.wml.STLineSpacingRule.AUTO;
import static org.docx4j.wml.STTblStyleOverrideType.FIRST_COL;
import static org.docx4j.wml.STTblStyleOverrideType.LAST_COL;
import static org.docx4j.wml.STVerticalJc.CENTER;

/**
 * @author sali
 */
public class CreateCustomStyle {

    @Test
    public void createAdmonitionTableStyle() {
        final PPrBuilder pPrBuilder = new PPrBuilder();
        Spacing spacing = pPrBuilder.getSpacingBuilder().withAfter(0L).withLine(240L).withLineRule(AUTO).getObject();
        PPr pPr = pPrBuilder.withSpacing(spacing).getObject();

        final CTTblPrBaseBuilder ctTblPrBaseBuilder = new CTTblPrBaseBuilder();
        TblBorders tblBorders = new TblBordersBuilder().withTop(getNilBorder()).withLeft(getNilBorder())
                .withBottom(getNilBorder()).withRight(getNilBorder()).withInsideH(getNilBorder()).withInsideV(getNilBorder())
                .getObject();
        CTTblPrBase tblPr = ctTblPrBaseBuilder.withTblBorders(tblBorders).getObject();

        CTTblStylePr[] ctTblStylePrs = new CTTblStylePr[2];

        // first col, "right" border and Vertical Align "Center"
        TcPrBuilder tcPrBuilder = new TcPrBuilder();
        TcPrInner.TcBorders tcBorders = tcPrBuilder.getTcBordersBuilder().withRight(getBorder(SINGLE, 8L, 0L, "DDDDD8")).getObject();
        TcPr tcPr = tcPrBuilder.withTcBorders(tcBorders).withVAlign(new CTVerticalJcBuilder().withVal(CENTER).getObject()).getObject();
        RPr rPr = new RPrBuilder().withColor(getColor("808080")).getObject();
        ctTblStylePrs[0] = new CTTblStylePrBuilder().withType(FIRST_COL).withRPr(rPr).withTcPr(tcPr).getObject();

        // second column, "left" border
        tcPrBuilder = new TcPrBuilder();
        tcBorders = tcPrBuilder.getTcBordersBuilder().withLeft(getBorder(SINGLE, 8L, 0L, "DDDDD8")).getObject();
        tcPr = tcPrBuilder.withTcBorders(tcBorders).getObject();
        rPr = new RPrBuilder().withColor(getColor("808080")).getObject();
        ctTblStylePrs[1] = new CTTblStylePrBuilder().withType(LAST_COL).withRPr(rPr).withTcPr(tcPr).getObject();

        StyleBuilder styleBuilder = new StyleBuilder().withType("table").withCustomStyle(true).withStyleId("AdmonitionTable")
                .withName("Admonition Table").withBasedOn("TableNormal").withUiPriority(47L).withRsid(nextId())
                .withPPr(pPr).withTblPr(tblPr).addTblStylePr(ctTblStylePrs);

        System.out.println(marshaltoString(styleBuilder.getObject()));
    }
}
