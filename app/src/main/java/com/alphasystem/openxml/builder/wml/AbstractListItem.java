package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RPr;

import static org.docx4j.wml.NumberFormat.DECIMAL;

/**
 * @author sali
 */
public abstract class AbstractListItem<T> implements ListItem<T> {

    private int numberId;
    private String name;
    private final String styleName;
    private final NumberFormat numberFormat;
    private final String id;

    protected AbstractListItem(String styleName, NumberFormat numberFormat, String id) {
        this.styleName = styleName;
        this.numberFormat = numberFormat;
        this.id = id;
        setName(styleName);
    }

    public AbstractListItem(String styleName, String id) {
        this(styleName, DECIMAL, id);
    }

    protected AbstractListItem(String styleName, NumberFormat numberFormat) {
        this(styleName, numberFormat, null);
    }

    @Override
    public void setNumberId(int numberId) {
        this.numberId = numberId;
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
        return false;
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
        return null;
    }

    @Override
    public long getLeftIndent(int level) {
        return 0;
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
        return "hybridMultilevel";
    }

    @Override
    public RPr getRPr() {
        return null;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public String getName() {
        return name;
    }
}
