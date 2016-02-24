package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RPr;

import java.util.LinkedHashMap;
import java.util.Map;

import static java.lang.String.format;

/**
 * @author sali
 */
public enum OrderedListItem implements ListItem<OrderedListItem> {

    ARABIC(1, "arabic", NumberFormat.DECIMAL, "04090019") {
        @Override
        public OrderedListItem getNext() {
            return LOWER_ALPHA;
        }
    },

    LOWER_ALPHA(2, "loweralpha", NumberFormat.LOWER_LETTER, "04090013") {
        @Override
        public OrderedListItem getNext() {
            return LOWER_ROMAN;
        }
    },

    LOWER_ROMAN(3, "lowerroman", NumberFormat.LOWER_ROMAN, "0409001B") {
        @Override
        public OrderedListItem getNext() {
            return UPPER_ALPHA;
        }
    },

    UPPER_ALPHA(4, "upperalpha", NumberFormat.UPPER_LETTER, "04090017") {
        @Override
        public OrderedListItem getNext() {
            return UPPER_ROMAN;
        }
    },

    UPPER_ROMAN(5, "upperroman", NumberFormat.UPPER_ROMAN, "0409000F") {
        @Override
        public OrderedListItem getNext() {
            return ARABIC;
        }
    };

    private static Map<String, OrderedListItem> valuesMap = new LinkedHashMap<>();

    static {
        for (OrderedListItem orderedListItem : values()) {
            valuesMap.put(orderedListItem.getStyleName(), orderedListItem);
        }
    }

    public static OrderedListItem getByStyleName(String styleName) {
        return valuesMap.get(styleName);
    }

    private final int numberId;
    private final String styleName;
    private final NumberFormat numberFormat;
    private final String id;

    OrderedListItem(int numberId, String styleName, NumberFormat numberFormat, String id) {
        this.numberId = numberId;
        this.styleName = styleName;
        this.numberFormat = numberFormat;
        this.id = id;
    }

    @Override
    public int getNumberId() {
        return numberId;
    }

    @Override
    public String getStyleName() {
        return styleName;
    }

    @Override
    public NumberFormat getNumberFormat() {
        return numberFormat;
    }

    @Override
    public String getId() {
        return id;
    }

    @Override
    public String getValue(int number) {
        return format("%%%s.", number);
    }

    @Override
    public RPr getRPr() {
        return null;
    }
}
