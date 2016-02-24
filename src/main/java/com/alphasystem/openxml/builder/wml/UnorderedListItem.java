package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

import java.util.LinkedHashMap;
import java.util.Map;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRPrBuilder;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.wml.NumberFormat.BULLET;

/**
 * @author sali
 */
public enum UnorderedListItem implements ListItem<UnorderedListItem> {

    DOT(6, "dot", "\u00B7", "04090001", R_FONTS_SYMBOL) {
        @Override
        public UnorderedListItem getNext() {
            return DIAMOND;
        }
    },

    DIAMOND(7, "diamond", "\u0076", "04090003", R_FONTS_WINDINGS) {
        @Override
        public UnorderedListItem getNext() {
            return SQUARE;
        }
    },

    SQUARE(8, "square", "\u00A7", "04090005", R_FONTS_WINDINGS) {
        @Override
        public UnorderedListItem getNext() {
            return CIRCLE;
        }
    },

    CIRCLE(9, "circle", "o", "04090007", R_FONTS_COURIER_NEW) {
        @Override
        public UnorderedListItem getNext() {
            return ARROW;
        }
    },

    ARROW(10, "arrow", "\u00D8", "04090009", R_FONTS_WINDINGS) {
        @Override
        public UnorderedListItem getNext() {
            return DOT;
        }
    };

    private static Map<String, UnorderedListItem> valuesMap = new LinkedHashMap<>();

    static {
        for (UnorderedListItem unorderedListItem : values()) {
            valuesMap.put(unorderedListItem.getStyleName(), unorderedListItem);
        }
    }

    public static UnorderedListItem getByStyleName(String styleName) {
        final UnorderedListItem item = isBlank(styleName) ? null : valuesMap.get(styleName);
        return (item == null) ? DOT : item;
    }

    private final int numberId;
    private final String styleName;
    private final String value;
    private final String id;
    private final NumberFormat numberFormat;
    private final RFonts font;

    UnorderedListItem(int numberId, String styleName, String value, String id, RFonts font) {
        this.numberId = numberId;
        this.styleName = styleName;
        this.value = value;
        this.id = id;
        this.font = font;
        this.numberFormat = BULLET;
    }

    @Override
    public int getNumberId() {
        return numberId;
    }

    @Override
    public String getStyleName() {
        return styleName;
    }

    public String getValue() {
        return value;
    }

    @Override
    public String getValue(int number) {
        return value;
    }

    @Override
    public String getId() {
        return id;
    }

    @Override
    public RPr getRPr() {
        return getRPrBuilder().withRFonts(font).getObject();
    }

    @Override
    public NumberFormat getNumberFormat() {
        return numberFormat;
    }
}
