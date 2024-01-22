package com.alphasystem.docx4j.builder.wml;

import org.docx4j.wml.NumberFormat;

import java.util.LinkedHashMap;
import java.util.Map;

import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;

/**
 * @author sali
 */
public class OrderedList extends AbstractListItem<OrderedList> {

    public static final OrderedList ARABIC = new OrderedList("arabic", NumberFormat.DECIMAL, "04090001");
    public static final OrderedList LOWER_ALPHA = new OrderedList("loweralpha", NumberFormat.LOWER_LETTER, "04090003");
    public static final OrderedList LOWER_ROMAN = new OrderedList("lowerroman", NumberFormat.LOWER_ROMAN, "04090005");
    public static final OrderedList UPPER_ALPHA = new OrderedList("upperalpha", NumberFormat.UPPER_LETTER, "04090007");
    public static final OrderedList UPPER_ROMAN = new OrderedList("upperroman", NumberFormat.UPPER_ROMAN, "04090009");

    private static final int LEFT_INDENT_VALUE = 720;
    private static final int HANGING_VALUE = 360;
    private static final Map<String, OrderedList> valuesMap = new LinkedHashMap<>();

    static {
        for (OrderedList item : values()) {
            valuesMap.put(item.getStyleName(), item);
        }
    }

    public static OrderedList getByStyleName(String styleName) {
        final OrderedList item = isBlank(styleName) ? null : valuesMap.get(styleName);
        return (item == null) ? ARABIC : item;
    }

    public static OrderedList[] values() {
        return new OrderedList[]{ARABIC, LOWER_ALPHA, LOWER_ROMAN, UPPER_ALPHA, UPPER_ROMAN};
    }

    OrderedList(String styleName, NumberFormat numberFormat, String id) {
        super(styleName, numberFormat, id);
    }

    @Override
    public String getValue(int number) {
        return format("%%%s.", number);
    }

    @Override
    public long getLeftIndent(int level) {
        return LEFT_INDENT_VALUE + ((long) LEFT_INDENT_VALUE * level);
    }

    @Override
    public long getHangingValue(int level) {
        return HANGING_VALUE;
    }

    @Override
    public Boolean isTentative(int level) {
        return level <= 0 ? null : Boolean.TRUE;
    }

}
