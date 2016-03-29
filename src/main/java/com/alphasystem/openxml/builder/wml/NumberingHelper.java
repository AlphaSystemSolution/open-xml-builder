package com.alphasystem.openxml.builder.wml;

import com.alphasystem.util.IdGenerator;
import org.docx4j.wml.*;
import org.docx4j.wml.Numbering.AbstractNum;

import java.util.ArrayList;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getCtLongHexNumber;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static org.apache.commons.lang3.ArrayUtils.add;

/**
 * @author sali
 */
public class NumberingHelper {

//    private static final String META_INF = "META-INF";
//    private static final String NUMBERING_FILE_NAME = "numbering.xml";

    @SuppressWarnings({"unchecked"})
    private static <T extends ListItem<T>> List<T> getListItems(T firstItem) {
        List<T> list = new ArrayList<>();
        list.add(firstItem);
        T currentItem = firstItem;
        while (true) {
            final T item = currentItem.getNext();
            if (item == null || item.equals(firstItem)) {
                break;
            }
            currentItem = item;
            list.add(item);
        }

        return list;
    }

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
        final LvlBuilder lvlBuilder = getLvlBuilder();
        NumFmt numFmt = getNumFmtBuilder().withVal(numFmtValue).getObject();
        return lvlBuilder.withIlvl(ilvl).withTplc(tplc).withTentative(tentative).withStart(startValue).withNumFmt(numFmt)
                .withLvlText(lvlTextValue, null).withPStyle(styleName).withLvlJc(jc).withPPr(pPr).withRPr(rPr).getObject();
    }

    private static PPr getPPr(long leftValue, long hangingValue) {
        final PPrBuilder pPrBuilder = getPPrBuilder();
        PPrBase.Ind ind = pPrBuilder.getIndBuilder().withLeft(leftValue).withHanging(hangingValue).getObject();
        return pPrBuilder.withInd(ind).getObject();
    }

    private static AbstractNum getAbstractNum(long id, String nsId, String tmpl, String multiLevel, Lvl[] lvls) {
        return getNumberingAbstractNumBuilder().withAbstractNumId(id).withNsid(getCtLongHexNumber(nsId))
                .withTmpl(getCtLongHexNumber(tmpl)).withMultiLevelType(multiLevel).addLvl(lvls).getObject();
    }

    private final NumberingBuilder numberingBuilder = getNumberingBuilder();
    private int numbers = 0;

    /*private void buildDefaultNumbering() throws IOException {
        Path path = save(get(USER_DIR, "src/main/resources", META_INF), );
        System.out.println(format("File created {%s}", path));

        path = save(get(USER_DIR, "src/main/resources", META_INF, "multi-level-heading"), getMultiLevelHeadingNumbering());
        System.out.println(format("File created {%s}", path));
    }*/

    public Numbering getNumbering() {
        return numberingBuilder.getObject();
    }

    public void populateDefaultNumbering() {
        for (OrderedList orderedList : OrderedList.values()) {
            populate(orderedList);
        }
        for (UnorderedList unorderedList : UnorderedList.values()) {
            populate(unorderedList);
        }
    }

//    public Numbering getMultiLevelHeadingNumbering() {
//        populate(HeadingList.HEADING1);
//        return numberingBuilder.getObject();
//    }

    public <T extends ListItem<T>> int populate(T item) {
        return populate(getListItems(item));
    }

    private <T extends ListItem<T>> int populate(final List<T> items) {
        numbers++;
        T firstItem = items.get(0);
        int numberId = firstItem.getNumberId();
        numberId = (numberId != numbers) ? numbers : numberId;
        long abstractNumId = numberId - 1;
        numberingBuilder.addAbstractNum(getAbstractNum(abstractNumId, IdGenerator.nextId(), IdGenerator.nextId(),
                firstItem.getMultiLevelType(), getLevels(items))).addNum(getNum(numberId));
        return numberId;
    }

//    public Path save(Path targetDir, Numbering numbering) throws IOException {
//        return write(get(targetDir.toString(), NUMBERING_FILE_NAME), marshaltoString(numbering).getBytes());
//    }

}
