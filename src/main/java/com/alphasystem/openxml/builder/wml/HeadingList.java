package com.alphasystem.openxml.builder.wml;

/**
 * @author sali
 */
public class HeadingList extends AbstractListItem<HeadingList> {

    public static final HeadingList HEADING1 = new HeadingList("ListHeading1", "List Heading 1", "04090021", "Heading1");

    public static final HeadingList HEADING2 = new HeadingList("ListHeading2", "List Heading 2", "Heading2");

    public static final HeadingList HEADING3 = new HeadingList("ListHeading3", "List Heading 3", "Heading3");

    public static final HeadingList HEADING4 = new HeadingList("ListHeading4", "List Heading 4", "Heading4");

    public static final HeadingList HEADING5 = new HeadingList("ListHeading5", "List Heading 5", "Heading5");

    private static final int LEFT_INDENT_VALUE = 432;
    private static final int INCREMENT_VALUE = 144;

    protected String baseStyle;
    protected boolean createNewStyle;

    public HeadingList(String styleName, String name, String baseStyle, boolean createNewStyle) {
        this(styleName, name, null, baseStyle, createNewStyle);
    }

    public HeadingList(String styleName, String name, String baseStyle) {
        this(styleName, name, null, baseStyle, true);
    }

    public HeadingList(String styleName, String name, String id, String baseStyle) {
        this(styleName, name, id, baseStyle, true);
    }

    public HeadingList(String styleName, String name, String id, String baseStyle, boolean createNewStyle) {
        super(styleName, id);
        setName(name);
        setBaseStyle(baseStyle);
        setCreateNewStyle(createNewStyle);
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

    public String getBaseStyle() {
        return baseStyle;
    }

    public void setBaseStyle(String baseStyle) {
        this.baseStyle = baseStyle;
    }

    public boolean isCreateNewStyle() {
        return createNewStyle;
    }

    public void setCreateNewStyle(boolean createNewStyle) {
        this.createNewStyle = createNewStyle;
    }
}
