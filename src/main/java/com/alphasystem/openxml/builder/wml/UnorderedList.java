package com.alphasystem.openxml.builder.wml;

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
public abstract class UnorderedList extends AbstractListItem<UnorderedList> {

    public static final UnorderedList DOT = new UnorderedList(6, "dot", "\u00B7", "04090011", R_FONTS_SYMBOL) {
        @Override
        public UnorderedList getNext() {
            return DIAMOND;
        }
    };
    public static final UnorderedList DIAMOND = new UnorderedList(7, "diamond", "\u0076", "04090013", R_FONTS_WINDINGS) {
        @Override
        public UnorderedList getNext() {
            return SQUARE;
        }
    };
    public static final UnorderedList SQUARE = new UnorderedList(8, "square", "\u00A7", "04090015", R_FONTS_WINDINGS) {
        @Override
        public UnorderedList getNext() {
            return CIRCLE;
        }
    };
    public static final UnorderedList CIRCLE = new UnorderedList(9, "circle", "o", "04090017", R_FONTS_COURIER_NEW) {
        @Override
        public UnorderedList getNext() {
            return ARROW;
        }
    };
    public static final UnorderedList ARROW = new UnorderedList(10, "arrow", "\u00D8", "04090019", R_FONTS_WINDINGS) {
        @Override
        public UnorderedList getNext() {
            return DOT;
        }
    };

    private static final int LEFT_INDENT_VALUE = 720;
    private static final int HANGING_VALUE = 360;

    private static Map<String, UnorderedList> valuesMap = new LinkedHashMap<>();

    static {
        for (UnorderedList item : values()) {
            valuesMap.put(item.getStyleName(), item);
        }
    }

    public static UnorderedList getByStyleName(String styleName) {
        final UnorderedList item = isBlank(styleName) ? null : valuesMap.get(styleName);
        return (item == null) ? DOT : item;
    }

    public static UnorderedList[] values() {
        return new UnorderedList[]{DOT, DIAMOND, SQUARE, CIRCLE, ARROW};
    }

    private final String value;
    private final RFonts font;

    UnorderedList(int numberId, String styleName, String value, String id, RFonts font) {
        super(numberId, styleName, BULLET, id);
        this.value = value;
        this.font = font;
    }

    @Override
    public String getValue(int number) {
        return value;
    }

    @Override
    public long getLeftIndent(int level) {
        return LEFT_INDENT_VALUE + (LEFT_INDENT_VALUE * level);
    }

    @Override
    public long getHangingValue(int level) {
        return HANGING_VALUE;
    }

    @Override
    public Boolean isTentative(int level) {
        return level <= 0 ? null : Boolean.TRUE;
    }

    @Override
    public RPr getRPr() {
        return getRPrBuilder().withRFonts(font).getObject();
    }

}
