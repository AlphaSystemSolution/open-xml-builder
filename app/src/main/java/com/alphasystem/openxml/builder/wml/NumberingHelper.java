package com.alphasystem.openxml.builder.wml;

import com.alphasystem.commons.util.IdGenerator;
import com.alphasystem.docx4j.builder.wml.NumberingBuilder;
import org.docx4j.wml.*;
import org.docx4j.wml.Numbering.AbstractNum;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicLong;

import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.openxml.builder.wml.OrderedList.*;
import static com.alphasystem.openxml.builder.wml.UnorderedList.*;
import static java.util.Arrays.asList;
import static org.apache.commons.lang3.ArrayUtils.add;

/**
 * @author sali
 */
public class NumberingHelper {

    private final Map<String, ListItem<?>> listItemsMap = new HashMap<>();

    private static Numbering.Num getNum(long numId) {
        return getNum(numId, numId - 1);
    }

    private static Numbering.Num getNum(long numId, long abstractNumIdValue) {
        return getNumberingNumBuilder().withNumId(numId).withAbstractNumId(abstractNumIdValue).getObject();
    }

    private static <T extends ListItem<T>> Lvl[] getLevels(List<T> listItems) {
        int level = 0;
        Lvl[] levels = new Lvl[0];
        for (T listItem : listItems) {
            levels = add(levels, getLevel(listItem, level));
            level++;
        }
        return levels;
    }

    private static <T extends ListItem<T>> Lvl getLevel(T item, int levelId) {
        final int number = levelId + 1;
        String levelTextValue = item.getValue(number);
        String styleName = item.linkStyle() ? item.getStyleName() : null;
        return getLevel(levelId, item.getId(), item.isTentative(levelId), item.getNumberFormat(), levelTextValue, styleName,
                getPPr(item.getLeftIndent(levelId), item.getHangingValue(levelId)), item.getRPr());
    }

    private static Lvl getLevel(long ilvl, String tplc, Boolean tentative, NumberFormat numFmtValue, String lvlTextValue,
                                String styleName, PPr pPr, RPr rPr) {
        return getLevel(ilvl, tplc, tentative, 1L, numFmtValue, lvlTextValue, styleName, JC_LEFT, pPr, rPr);
    }

    private static Lvl getLevel(long ilvl, String tplc, Boolean tentative, long startValue, NumberFormat numFmtValue,
                                String lvlTextValue, String styleName, Jc jc, PPr pPr, RPr rPr) {
        final var lvlBuilder = getLvlBuilder();
        NumFmt numFmt = getNumFmtBuilder().withVal(numFmtValue).getObject();
        return lvlBuilder.withIlvl(ilvl).withTplc(tplc).withTentative(tentative).withStart(startValue).withNumFmt(numFmt)
                .withLvlText(lvlTextValue, null).withPStyle(styleName).withLvlJc(jc).withPPr(pPr).withRPr(rPr).getObject();
    }

    private static PPr getPPr(long leftValue, long hangingValue) {
        final var pPrBuilder = getPPrBuilder();
        PPrBase.Ind ind = pPrBuilder.getIndBuilder().withLeft(leftValue).withHanging(hangingValue).getObject();
        return pPrBuilder.withInd(ind).getObject();
    }

    private static AbstractNum getAbstractNum(long id, String nsId, String tmpl, String multiLevel, Lvl[] lvls) {
        return getNumberingAbstractNumBuilder().withAbstractNumId(id).withNsid(nsId).withTmpl(tmpl)
                .withMultiLevelType(multiLevel).addLvl(lvls).getObject();
    }

    private static final NumberingHelper instance;

    static {
        instance = new NumberingHelper();
        instance.populateDefaultNumbering();
    }

    public static NumberingHelper getInstance() {
        return instance;
    }

    private final NumberingBuilder numberingBuilder = getNumberingBuilder();
    private final AtomicLong numberId = new AtomicLong(0);

    private NumberingHelper(){
    }

    public Numbering getNumbering() {
        return numberingBuilder.getObject();
    }

    public ListItem<?> getListItem(String styleName) {
        return listItemsMap.get(styleName);
    }

    private void populateDefaultNumbering() {
        populate(ARABIC, LOWER_ALPHA, LOWER_ROMAN, UPPER_ALPHA, UPPER_ROMAN);
        populate(LOWER_ALPHA, LOWER_ROMAN, UPPER_ALPHA, UPPER_ROMAN, ARABIC);
        populate(LOWER_ROMAN, UPPER_ALPHA, UPPER_ROMAN, ARABIC, LOWER_ALPHA);
        populate(UPPER_ALPHA, UPPER_ROMAN, ARABIC, LOWER_ALPHA, LOWER_ROMAN);
        populate(UPPER_ROMAN, ARABIC, LOWER_ALPHA, LOWER_ROMAN, UPPER_ALPHA);

        populate(DOT, DIAMOND, SQUARE, CIRCLE, ARROW);
        populate(DIAMOND, SQUARE, CIRCLE, ARROW, DOT);
        populate(SQUARE, CIRCLE, ARROW, DOT, DIAMOND);
        populate(CIRCLE, ARROW, DOT, DIAMOND, SQUARE);
        populate(ARROW, DOT, DIAMOND, SQUARE, CIRCLE);
    }

    @SafeVarargs
    public final <T extends ListItem<T>> long populate(T... items) {
        T firstItem = items[0];
        var numberId = this.numberId.addAndGet(1);
        firstItem.setNumberId(numberId);
        long abstractNumId = numberId - 1;
        numberingBuilder.addAbstractNum(getAbstractNum(abstractNumId, IdGenerator.nextId(), IdGenerator.nextId(),
                firstItem.getMultiLevelType(), getLevels(asList(items)))).addNum(getNum(numberId));
        listItemsMap.put(firstItem.getStyleName(), firstItem);
        return numberId;
    }

    public void update(String styleName, long numId) {
        final var listItem = listItemsMap.get(styleName);
        if (!Objects.isNull(listItem)) {
            listItem.setNumberId(numId);
            listItemsMap.put(styleName, listItem);
        }
    }

}
