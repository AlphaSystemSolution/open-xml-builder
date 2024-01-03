package com.alphasystem.openxml.builder.wml;

/**
 * @author sali
 */
public class HeadingList<T extends HeadingList<T>> extends AbstractListItem<T> {

    private static final int LEFT_INDENT_VALUE = 432;
    private static final int INCREMENT_VALUE = 144;

    public HeadingList(String styleName) {
        this(styleName, null);
    }

    public HeadingList(String styleName, String id) {
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
        return LEFT_INDENT_VALUE + ((long) INCREMENT_VALUE * level);
    }

    @Override
    public String getMultiLevelType() {
        return "multilevel";
    }

}
