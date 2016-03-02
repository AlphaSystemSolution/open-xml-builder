package com.alphasystem.openxml.builder.wml;

import com.alphasystem.util.IdGenerator;
import org.docx4j.wml.*;
import org.docx4j.wml.Numbering.AbstractNum;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.getCtLongHexNumber;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.nio.NIOFileUtils.USER_DIR;
import static java.nio.file.Files.write;
import static java.nio.file.Paths.get;
import static org.apache.commons.lang3.ArrayUtils.add;
import static org.docx4j.XmlUtils.marshaltoString;

/**
 * @author sali
 */
public class NumberingHelper {

    private static final String META_INF = "META-INF";
    private static final String NUMBERING_FILE_NAME = "numbering.xml";

    public static void main(String[] args) {
        try {
            buildDefaultNumbering();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void buildDefaultNumbering() throws IOException {
        Path path = save(get(USER_DIR, "src/main/resources", META_INF), getDefaultNumbering());
        System.out.println(String.format("File created {%s}", path));

        path = save(get(USER_DIR, "src/main/resources", META_INF, "multi-level-heading"), getMultiLevelHeadingNumbering());
        System.out.println(String.format("File created {%s}", path));
    }

    public static Numbering getDefaultNumbering() {
        final NumberingBuilder numberingBuilder = getNumberingBuilder();
        populate(numberingBuilder, OrderedList.values());
        populate(numberingBuilder, UnorderedList.values());
        return numberingBuilder.getObject();
    }

    public static Numbering getMultiLevelHeadingNumbering() {
        final NumberingBuilder numberingBuilder = getNumberingBuilder();
        populate(numberingBuilder, HeadingList.HEADING1);
        return numberingBuilder.getObject();
    }

    @SafeVarargs
    public static <T extends ListItem<T>> void populate(final NumberingBuilder numberingBuilder, T... items) {
        for (T item : items) {
            populate(numberingBuilder, getListItems(item));
        }
    }

    public static <T extends ListItem<T>> void populate(NumberingBuilder numberingBuilder, final List<T> items) {
        T firstItem = items.get(0);
        long abstractNumId = firstItem.getNumberId() - 1;
        numberingBuilder.addAbstractNum(getAbstractNum(abstractNumId, IdGenerator.nextId(), IdGenerator.nextId(),
                firstItem.getMultiLevelType(), getLevels(items))).addNum(getNum(firstItem.getNumberId()));
    }

    public static Path save(Path targetDir, Numbering numbering) throws IOException {
        return write(get(targetDir.toString(), NUMBERING_FILE_NAME), marshaltoString(numbering).getBytes());
    }

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

    private static AbstractNum getAbstractNum(long id, String nsId, String tmpl, String multiLevel, Lvl[] lvls) {
        return getNumberingAbstractNumBuilder().withAbstractNumId(id).withNsid(getCtLongHexNumber(nsId))
                .withTmpl(getCtLongHexNumber(tmpl)).withMultiLevelType(multiLevel).addLvl(lvls).getObject();
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

}
