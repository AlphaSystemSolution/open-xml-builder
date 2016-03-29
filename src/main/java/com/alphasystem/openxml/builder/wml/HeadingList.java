package com.alphasystem.openxml.builder.wml;

/**
 * @author sali
 */
public class HeadingList extends AbstractListItem<HeadingList> {

    public static final HeadingList HEADING1 = new HeadingList("ListHeading1", "04090021") {
        @Override
        public String getName() {
            return "HEADING1";
        }
    };

    public static final HeadingList HEADING2 = new HeadingList("ListHeading2") {
        @Override
        public String getName() {
            return "HEADING2";
        }
    };

    public static final HeadingList HEADING3 = new HeadingList("ListHeading3") {
        @Override
        public String getName() {
            return "HEADING3";
        }
    };

    public static final HeadingList HEADING4 = new HeadingList("ListHeading4") {
        @Override
        public String getName() {
            return "HEADING4";
        }
    };

    public static final HeadingList HEADING5 = new HeadingList("ListHeading5") {
        @Override
        public String getName() {
            return "HEADING5";
        }
    };


    private static final int LEFT_INDENT_VALUE = 432;
    private static final int INCREMENT_VALUE = 144;

    HeadingList(String styleName) {
        this(styleName, null);
    }

    HeadingList(String styleName, String id) {
        super(styleName, id);
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
    public boolean linkStyle() {
        return true;
    }

    @Override
    public long getLeftIndent(int level) {
        return LEFT_INDENT_VALUE + (INCREMENT_VALUE * level);
    }

    @Override
    public String getMultiLevelType() {
        return "multilevel";
    }
}
