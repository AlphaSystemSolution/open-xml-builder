package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RPr;

import static org.docx4j.wml.NumberFormat.DECIMAL;

/**
 * @author sali
 */
public enum HeadingListItem implements ListItem<HeadingListItem> {

    HEADING1(11, "Heading1", "04090021") {
        @Override
        public HeadingListItem getNext() {
            return HEADING2;
        }
    },

    HEADING2(-1, "Heading2", "04090023") {
        @Override
        public HeadingListItem getNext() {
            return HEADING3;
        }
    },

    HEADING3(-1, "Heading3", "04090025") {
        @Override
        public HeadingListItem getNext() {
            return HEADING4;
        }
    },

    HEADING4(-1, "Heading4", "04090027") {
        @Override
        public HeadingListItem getNext() {
            return HEADING5;
        }
    },

    HEADING5(-1, "Heading5", "04090029") {
        @Override
        public HeadingListItem getNext() {
            return HEADING1;
        }
    };

    private static final int LEFT_INDENT_VALUE = 432;
    private static final int INCREMENT_VALUE = 144;

    private final int numberId;
    private final String styleName;
    private final NumberFormat numberFormat;
    private final String id;

    HeadingListItem(int numberId, String styleName, String id) {
        this.numberId = numberId;
        this.styleName = styleName;
        this.numberFormat = DECIMAL;
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
    public boolean linkStyle() {
        return true;
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
        StringBuilder builder = new StringBuilder();
        for (int i = 1; i <= number; i++) {
            builder.append("%").append(i).append(".");
        }
        return builder.toString();
    }

    @Override
    public long getLeftIndent(int level) {
        return LEFT_INDENT_VALUE + (INCREMENT_VALUE * level);
    }

    @Override
    public long getHangingValue(int level) {
        return getLeftIndent(level);
    }

    @Override
    public Boolean isTentative(int level) {
        return null;
    }

    @Override
    public String getMultiLevelType() {
        return "multilevel";
    }

    @Override
    public RPr getRPr() {
        return null;
    }

}
