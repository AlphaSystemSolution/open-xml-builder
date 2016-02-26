package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RPr;

import static org.docx4j.wml.NumberFormat.DECIMAL;

/**
 * @author sali
 */
public abstract class AbstractListItem<T> implements ListItem<T> {

    private final int numberId;
    private final String styleName;
    private final NumberFormat numberFormat;
    private final String id;

    protected AbstractListItem(int numberId, String styleName, NumberFormat numberFormat, String id) {
        this.numberId = numberId;
        this.styleName = styleName;
        this.numberFormat = numberFormat;
        this.id = id;
    }

    public AbstractListItem(int numberId, String styleName, String id) {
        this.numberId = numberId;
        this.styleName = styleName;
        this.numberFormat = DECIMAL;
        this.id = id;
    }

    protected AbstractListItem(int numberId, String styleName, NumberFormat numberFormat) {
        this(numberId, styleName, numberFormat, null);
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

}
