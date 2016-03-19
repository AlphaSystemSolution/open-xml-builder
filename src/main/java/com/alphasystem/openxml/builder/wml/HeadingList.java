package com.alphasystem.openxml.builder.wml;

/**
 * @author sali
 */
public abstract class HeadingList extends AbstractListItem<HeadingList> {

    public static final HeadingList HEADING5 = new HeadingList("ListHeading5") {
        @Override
        public HeadingList getNext() {
            return null;
        }

        @Override
        public String getName() {
            return "HEADING5";
        }
    };
    public static final HeadingList HEADING4 = new HeadingList("ListHeading4") {
        @Override
        public HeadingList getNext() {
            return HEADING5;
        }

        @Override
        public String getName() {
            return "HEADING4";
        }
    };
    public static final HeadingList HEADING3 = new HeadingList("ListHeading3") {
        @Override
        public HeadingList getNext() {
            return HEADING4;
        }

        @Override
        public String getName() {
            return "HEADING3";
        }
    };
    public static final HeadingList HEADING2 = new HeadingList("ListHeading2") {
        @Override
        public HeadingList getNext() {
            return HEADING3;
        }

        @Override
        public String getName() {
            return "HEADING2";
        }
    };
    public static final HeadingList HEADING1 = new HeadingList(11, "ListHeading1", "04090021") {
        @Override
        public HeadingList getNext() {
            return HEADING2;
        }

        @Override
        public String getName() {
            return "HEADING1";
        }
    };
    private static final int LEFT_INDENT_VALUE = 432;
    private static final int INCREMENT_VALUE = 144;

    HeadingList(String styleName) {
        this(-1, styleName, null);
    }

    HeadingList(int numberId, String styleName, String id) {
        super(numberId, styleName, id);
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
