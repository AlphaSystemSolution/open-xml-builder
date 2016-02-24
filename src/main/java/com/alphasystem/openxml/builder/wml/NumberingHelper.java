package com.alphasystem.openxml.builder.wml;

import com.alphasystem.util.IdGenerator;
import org.docx4j.wml.*;
import org.docx4j.wml.Numbering.AbstractNum;
import org.docx4j.wml.Numbering.AbstractNum.MultiLevelType;

import java.util.ArrayList;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getCtLongHexNumber;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static java.lang.Boolean.TRUE;
import static org.apache.commons.lang3.ArrayUtils.add;

/**
 * @author sali
 */
public class NumberingHelper {

    private static final int LEFT_INDENT_VALUE = 720;
    private static final int HANGING_VALUE = 360;

    public static Numbering createNumbering() {
        final NumberingBuilder numberingBuilder = getNumberingBuilder();
        populate(numberingBuilder);
        return numberingBuilder.getObject();
    }

    @SuppressWarnings({"unchecked"})
    private static <T extends Enum<T> & ListItem<T>> List<T> getListItems(T firstItem) {
        List<T> list = new ArrayList<>();
        list.add(firstItem);
        T currentItem = firstItem;
        while (true) {
            final T item = currentItem.getNext();
            if (item.equals(firstItem)) {
                break;
            }
            currentItem = item;
            list.add(item);
        }

        return list;
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

    private static void populate(NumberingBuilder numberingBuilder) {
        for (OrderedListItem orderedListItem : OrderedListItem.values()) {
            populateOrderedListItem(numberingBuilder, orderedListItem);
        }
        for (UnorderedListItem unorderedListItem : UnorderedListItem.values()) {
            populateUnorderedListItem(numberingBuilder, unorderedListItem);
        }
    }

    private static void populateOrderedListItem(NumberingBuilder numberingBuilder, OrderedListItem initialItem) {
        final List<OrderedListItem> orderedListItems = getListItems(initialItem);
        long abstractNumId = initialItem.getNumberId() - 1;
        numberingBuilder.addAbstractNum(getAbstractNum(abstractNumId, IdGenerator.nextId(), IdGenerator.nextId(),
                getOrderedListLevels(orderedListItems))).addNum(getNum(initialItem.getNumberId()));
    }

    private static void populateUnorderedListItem(NumberingBuilder numberingBuilder, UnorderedListItem initialItem) {
        final List<UnorderedListItem> unorderedListItems = getListItems(initialItem);
        long abstractNumId = initialItem.getNumberId() - 1;
        numberingBuilder.addAbstractNum(getAbstractNum(abstractNumId, IdGenerator.nextId(), IdGenerator.nextId(),
                getUnorderedListLevels(unorderedListItems))).addNum(getNum(initialItem.getNumberId()));
    }

    private static Lvl[] getOrderedListLevels(List<OrderedListItem> orderedListItems) {
        int level = 0;
        Lvl[] lvls = new Lvl[0];
        for (OrderedListItem orderedListItem : orderedListItems) {
            lvls = add(lvls, getLevel(orderedListItem, level));
            level++;
        }
        return lvls;
    }

    private static Lvl[] getUnorderedListLevels(List<UnorderedListItem> unorderedListItems) {
        int level = 0;
        Lvl[] lvls = new Lvl[0];
        for (UnorderedListItem unorderedListItem : unorderedListItems) {
            lvls = add(lvls, getLevel(unorderedListItem, level));
            level++;
        }
        return lvls;
    }

    private static <T extends Enum<T> & ListItem<T>> Lvl getLevel(T item, int levelId) {
        final boolean initialLevel = levelId <= 0L;
        Boolean tentative = initialLevel ? null : TRUE;
        final int number = levelId + 1;
        String levelTextValue = item.getValue(number);
        long leftIndentValue = initialLevel ? LEFT_INDENT_VALUE : (LEFT_INDENT_VALUE * number);
        return getLvl(levelId, item.getId(), tentative, item.getNumberFormat(), levelTextValue,
                getPPr(leftIndentValue, HANGING_VALUE), item.getRPr());
    }

    private static Lvl getLvl(long ilvl, String tplc, Boolean tentative, NumberFormat numFmtValue, String lvlTextValue,
                              PPr pPr, RPr rPr) {
        return getLvl(ilvl, tplc, tentative, 1L, numFmtValue, lvlTextValue, JC_LEFT, pPr, rPr);
    }

    private static Lvl getLvl(long ilvl, String tplc, Boolean tentative, long startValue, NumberFormat numFmtValue,
                              String lvlTextValue, Jc jc, PPr pPr, RPr rPr) {
        Lvl.Start start = getLvlStartBuilder().withVal(startValue).getObject();
        NumFmt numFmt = getNumFmtBuilder().withVal(numFmtValue).getObject();
        Lvl.LvlText lvlText = getLvlLvlTextBuilder().withVal(lvlTextValue).getObject();
        return getLvlBuilder().withIlvl(ilvl).withTplc(tplc).withTentative(tentative).withStart(start).withNumFmt(numFmt)
                .withLvlText(lvlText).withLvlJc(jc).withPPr(pPr).withRPr(rPr).getObject();
    }

    private static PPr getPPr(long leftValue, long hangingValue) {
        PPrBase.Ind ind = getPPrBaseIndBuilder().withLeft(leftValue).withHanging(hangingValue).getObject();
        return getPPrBuilder().withInd(ind).getObject();
    }

}
