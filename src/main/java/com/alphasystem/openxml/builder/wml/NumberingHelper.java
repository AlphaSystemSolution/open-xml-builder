package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.*;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.AbstractNum.MultiLevelType;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getCtLongHexNumber;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static java.lang.Boolean.TRUE;
import static org.docx4j.wml.NumberFormat.*;
import static org.docx4j.wml.STHint.DEFAULT;

/**
 * @author sali
 */
public class NumberingHelper {

    private static final RFonts R_FONTS_WINDINGS = getRFonts("Wingdings", "Wingdings");
    public static final RFonts R_FONTS_COURIER_NEW = getRFonts("Courier New", "Courier New", "Courier New");
    public static final RFonts R_FONTS_SYMBOL = getRFonts("Symbol", "Symbol");

    public static Numbering createNumbering() {
        return getNumberingBuilder().addAbstractNum(getAbstractNum0(), getAbstractNum1(), getAbstractNum2(),
                getAbstractNum3(), getAbstractNum4(), getAbstractNum5(), getAbstractNum6(), getAbstractNum7(),
                getAbstractNum8(), getAbstractNum9()).addNum(getNum(1), getNum(2), getNum(3), getNum(4), getNum(5),
                getNum(6), getNum(7), getNum(8), getNum(9), getNum(10)).getObject();
    }

    private static AbstractNum getAbstractNum(long id, String nsId, String tmpl, Lvl[] lvls) {
        final MultiLevelType multiLevelType = getNumberingAbstractNumMultiLevelTypeBuilder()
                .withVal("hybridMultilevel").getObject();
        return getNumberingAbstractNumBuilder().withAbstractNumId(id).withNsid(getCtLongHexNumber(nsId))
                .withTmpl(getCtLongHexNumber(tmpl)).withMultiLevelType(multiLevelType).addLvl(lvls).getObject();
    }

    private static Numbering.Num getNum(long numId, long abstractNumIdValue) {
        final Numbering.Num.AbstractNumId abstractNumId = getNumberingNumAbstractNumIdBuilder().withVal(abstractNumIdValue).getObject();
        return getNumberingNumBuilder().withNumId(numId).withAbstractNumId(abstractNumId).getObject();
    }

    private static Numbering.Num getNum(long numId) {
        return getNum(numId, numId - 1);
    }

    private static AbstractNum getAbstractNum0() {
        return getAbstractNum(0L, "4BAB24DE", "2FD0B640", getLvl0());
    }

    private static AbstractNum getAbstractNum1() {
        return getAbstractNum(1L, "269A1BDF", "C80AC14C", getLvl1());
    }

    private static AbstractNum getAbstractNum2() {
        return getAbstractNum(2L, "6AA214B3", "84B0FCEA", getLvl2());
    }

    private static AbstractNum getAbstractNum3() {
        return getAbstractNum(3L, "014C68C2", "2A206E70", getLvl3());
    }

    private static AbstractNum getAbstractNum4() {
        return getAbstractNum(4L, "3DD8374A", "B868EB5A", getLvl4());
    }

    private static AbstractNum getAbstractNum5() {
        return getAbstractNum(5L, "211A17A4", "2C6481B8", getLvl5());
    }

    private static AbstractNum getAbstractNum6() {
        return getAbstractNum(6L, "22FF4E64", "37D2C838", getLvl6());
    }

    private static AbstractNum getAbstractNum7() {
        return getAbstractNum(7L, "34611B3C", "21A2C23C", getLvl7());
    }

    private static AbstractNum getAbstractNum8() {
        return getAbstractNum(8L, "0E807AB3", "FD4E2FB8", getLvl8());
    }

    private static AbstractNum getAbstractNum9() {
        return getAbstractNum(9L, "03224B5A", "0AAE169E", getLvl9());
    }


    private static Lvl[] getLvl0() {
        return new Lvl[]{
                getLvl(0L, "DAEC4FFA", null, 1L, DECIMAL, "%1.", JC_LEFT, getPPr(720, 360)),
                getLvl(1L, "04090019", TRUE, 1L, UPPER_ROMAN, "%2.", JC_LEFT, getPPr(1440, 360)),
                getLvl(2L, "0409001B", TRUE, 1L, UPPER_LETTER, "%3.", JC_RIGHT, getPPr(2160, 180)),
                getLvl(3L, "0409000F", TRUE, 1L, LOWER_ROMAN, "%4.", JC_LEFT, getPPr(2880, 360)),
                getLvl(4L, "04090019", TRUE, 1L, LOWER_LETTER, "%5.", JC_LEFT, getPPr(3600, 360))};
    }

    private static Lvl[] getLvl1() {
        return new Lvl[]{
                getLvl(0L, "04090013", null, 1L, UPPER_ROMAN, "%1.", JC_LEFT, getPPr(720, 360)),
                getLvl(1L, "04090019", TRUE, 1L, UPPER_LETTER, "%2.", JC_LEFT, getPPr(1440, 360)),
                getLvl(2L, "0409001B", TRUE, 1L, LOWER_ROMAN, "%3.", JC_RIGHT, getPPr(2160, 180)),
                getLvl(3L, "0409000F", TRUE, 1L, LOWER_LETTER, "%4.", JC_LEFT, getPPr(2880, 360)),
                getLvl(4L, "04090019", TRUE, 1L, DECIMAL, "%5.", JC_LEFT, getPPr(3600, 360))};
    }

    private static Lvl[] getLvl2() {
        return new Lvl[]{
                getLvl(0L, "04090017", null, 1L, UPPER_LETTER, "%1.", JC_LEFT, getPPr(720, 360)),
                getLvl(1L, "04090019", TRUE, 1L, LOWER_ROMAN, "%2.", JC_LEFT, getPPr(1440, 360)),
                getLvl(2L, "0409001B", TRUE, 1L, LOWER_LETTER, "%3.", JC_RIGHT, getPPr(2160, 180)),
                getLvl(3L, "0409000F", TRUE, 1L, DECIMAL, "%4.", JC_LEFT, getPPr(2880, 360)),
                getLvl(4L, "04090019", TRUE, 1L, UPPER_ROMAN, "%5.", JC_LEFT, getPPr(3600, 360))};
    }

    private static Lvl[] getLvl3() {
        return new Lvl[]{
                getLvl(0L, "0409001B", null, 1L, LOWER_ROMAN, "%1.", JC_RIGHT, getPPr(720, 360)),
                getLvl(1L, "04090019", TRUE, 1L, LOWER_LETTER, "%2.", JC_LEFT, getPPr(1440, 360)),
                getLvl(2L, "0409001B", TRUE, 1L, DECIMAL, "%3.", JC_RIGHT, getPPr(2160, 180)),
                getLvl(3L, "0409000F", TRUE, 1L, UPPER_ROMAN, "%4.", JC_LEFT, getPPr(2880, 360)),
                getLvl(4L, "04090019", TRUE, 1L, UPPER_LETTER, "%5.", JC_LEFT, getPPr(3600, 360))};
    }

    private static Lvl[] getLvl4() {
        return new Lvl[]{
                getLvl(0L, "04090019", null, 1L, LOWER_LETTER, "%1.", JC_LEFT, getPPr(720, 360)),
                getLvl(1L, "04090019", TRUE, 1L, DECIMAL, "%2.", JC_LEFT, getPPr(1440, 360)),
                getLvl(2L, "0409001B", TRUE, 1L, UPPER_ROMAN, "%3.", JC_RIGHT, getPPr(2160, 180)),
                getLvl(3L, "0409000F", TRUE, 1L, UPPER_LETTER, "%4.", JC_LEFT, getPPr(2880, 360)),
                getLvl(4L, "04090019", TRUE, 1L, LOWER_ROMAN, "%5.", JC_LEFT, getPPr(3600, 360))};
    }

    private static Lvl[] getLvl5() {
        return new Lvl[]{
                getLvl(0L, "04090001", null, 1L, BULLET, "\uF0B7", JC_LEFT, getPPr(720, 360), getRPr(R_FONTS_SYMBOL)),
                getLvl(1L, "04090003", TRUE, 1L, BULLET, "\u00A7", JC_LEFT, getPPr(1440, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(2L, "04090005", TRUE, 1L, BULLET, "o", JC_LEFT, getPPr(2160, 360), getRPr(R_FONTS_COURIER_NEW)),
                getLvl(3L, "04090001", TRUE, 1L, BULLET, "\uF0D8", JC_LEFT, getPPr(2880, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(4L, "04090003", TRUE, 1L, BULLET, "\uF0FC", JC_LEFT, getPPr(3600, 360), getRPr(R_FONTS_WINDINGS))};
    }

    private static Lvl[] getLvl6() {
        return new Lvl[]{
                getLvl(0L, "04090005", null, 1L, BULLET, "\u00A7", JC_LEFT, getPPr(720, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(1L, "04090003", TRUE, 1L, BULLET, "o", JC_LEFT, getPPr(1440, 360), getRPr(R_FONTS_COURIER_NEW)),
                getLvl(2L, "04090005", TRUE, 1L, BULLET, "\uF0D8", JC_LEFT, getPPr(2160, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(3L, "04090001", TRUE, 1L, BULLET, "\uF0FC", JC_LEFT, getPPr(2880, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(4L, "04090003", TRUE, 1L, BULLET, "\uF0B7", JC_LEFT, getPPr(3600, 360), getRPr(R_FONTS_SYMBOL))};
    }

    private static Lvl[] getLvl7() {
        return new Lvl[]{
                getLvl(0L, "04090003", null, 1L, BULLET, "o", JC_LEFT, getPPr(720, 360), getRPr(R_FONTS_COURIER_NEW)),
                getLvl(1L, "04090003", TRUE, 1L, BULLET, "\uF0D8", JC_LEFT, getPPr(1440, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(2L, "04090005", TRUE, 1L, BULLET, "\uF0FC", JC_LEFT, getPPr(2160, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(3L, "04090001", TRUE, 1L, BULLET, "\uF0B7", JC_LEFT, getPPr(2880, 360), getRPr(R_FONTS_SYMBOL)),
                getLvl(4L, "04090003", TRUE, 1L, BULLET, "\u00A7", JC_LEFT, getPPr(3600, 360), getRPr(R_FONTS_WINDINGS))};
    }

    private static Lvl[] getLvl8() {
        return new Lvl[]{
                getLvl(0L, "0409000B", null, 1L, BULLET, "\uF0D8", JC_LEFT, getPPr(720, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(1L, "04090003", TRUE, 1L, BULLET, "\uF0FC", JC_LEFT, getPPr(1440, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(2L, "04090005", TRUE, 1L, BULLET, "\uF0B7", JC_LEFT, getPPr(2160, 360), getRPr(R_FONTS_SYMBOL)),
                getLvl(3L, "04090001", TRUE, 1L, BULLET, "\u00A7", JC_LEFT, getPPr(2880, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(4L, "04090003", TRUE, 1L, BULLET, "o", JC_LEFT, getPPr(3600, 360), getRPr(R_FONTS_COURIER_NEW))};
    }

    private static Lvl[] getLvl9() {
        return new Lvl[]{
                getLvl(0L, "0409000D", null, 1L, BULLET, "\uF0FC", JC_LEFT, getPPr(720, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(1L, "04090003", TRUE, 1L, BULLET, "\uF0B7", JC_LEFT, getPPr(1440, 360), getRPr(R_FONTS_SYMBOL)),
                getLvl(2L, "0409000D", TRUE, 1L, BULLET, "\u00A7", JC_LEFT, getPPr(2160, 360), getRPr(R_FONTS_WINDINGS)),
                getLvl(3L, "04090001", TRUE, 1L, BULLET, "o", JC_LEFT, getPPr(2880, 360), getRPr(R_FONTS_COURIER_NEW)),
                getLvl(4L, "04090003", TRUE, 1L, BULLET, "\uF0D8", JC_LEFT, getPPr(3600, 360), getRPr(R_FONTS_WINDINGS))};
    }

    private static Lvl getLvl(long ilvl, String tplc, Boolean tentative, long startValue, NumberFormat numFmtValue,
                              String lvlTextValue, Jc jc, PPr pPr) {
        return getLvl(ilvl, tplc, tentative, startValue, numFmtValue, lvlTextValue, jc, pPr, null);
    }

    private static Lvl getLvl(long ilvl, String tplc, Boolean tentative, long startValue, NumberFormat numFmtValue,
                              String lvlTextValue, Jc jc, PPr pPr, RPr rPr) {
        return getLvl(ilvl, tplc, tentative, startValue, numFmtValue, lvlTextValue, null, jc, pPr, rPr);
    }

    private static Lvl getLvl(long ilvl, String tplc, Boolean tentative, long startValue, NumberFormat numFmtValue,
                              String lvlTextValue, Long picBulletIdValue, Jc jc, PPr pPr, RPr rPr) {
        Lvl.Start start = getLvlStartBuilder().withVal(startValue).getObject();
        NumFmt numFmt = getNumFmtBuilder().withVal(numFmtValue).getObject();
        Lvl.LvlText lvlText = getLvlLvlTextBuilder().withVal(lvlTextValue).getObject();
        Lvl.LvlPicBulletId picBulletId = null;
        if (picBulletIdValue != null) {
            picBulletId = getLvlLvlPicBulletIdBuilder().withVal(picBulletIdValue).getObject();
        }
        return getLvlBuilder().withIlvl(ilvl).withTplc(tplc).withTentative(tentative).withStart(start).withNumFmt(numFmt)
                .withLvlText(lvlText).withLvlJc(jc).withPPr(pPr).withRPr(rPr).withLvlPicBulletId(picBulletId).getObject();
    }

    private static PPr getPPr(long leftValue, long hangingValue) {
        PPrBase.Ind ind = getPPrBaseIndBuilder().withLeft(leftValue).withHanging(hangingValue).getObject();
        return getPPrBuilder().withInd(ind).getObject();
    }

    private static RPr getRPr(RFonts rFonts) {
        return getRPrBuilder().withRFonts(rFonts).getObject();
    }

    private static RFonts getRFonts(String ascii, String hAnsi) {
        return getRFonts(ascii, hAnsi, null);
    }

    private static RFonts getRFonts(String ascii, String hAnsi, String cs) {
        return getRFontsBuilder().withAscii(ascii).withHAnsi(hAnsi).withCs(cs).withHint(DEFAULT).getObject();
    }
}
