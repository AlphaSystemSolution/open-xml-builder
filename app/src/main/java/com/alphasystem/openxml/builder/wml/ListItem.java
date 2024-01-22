package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRFontsBuilder;
import static org.docx4j.wml.STHint.DEFAULT;

/**
 * Interface for common properties of a list in document.
 *
 * @author sali
 */
public interface ListItem<T> {

    RFonts R_FONTS_COURIER_NEW = getRFonts("Courier New", "Courier New", "Courier New");
    RFonts R_FONTS_SYMBOL = getRFonts("Symbol", "Symbol");
    RFonts R_FONTS_WINDINGS = getRFonts("Wingdings", "Wingdings");

    static RFonts getRFonts(String ascii, String hAnsi) {
        return getRFonts(ascii, hAnsi, null);
    }

    static RFonts getRFonts(String ascii, String hAnsi, String cs) {
        return getRFontsBuilder().withAscii(ascii).withHAnsi(hAnsi).withCs(cs).withHint(DEFAULT).getObject();
    }

    /**
     * Sets the number id for this item.
     *
     * @param numberId current number Id
     */
    void setNumberId(long numberId);

    /**
     * ID of the list.
     *
     * @return list id
     */
    long getNumberId();

    /**
     * Name of the style.
     *
     * @return style name
     */
    String getStyleName();

    /**
     * Whether or not link style with numbering system
     *
     * @return true if style name to be linked, false otherwise
     */
    boolean linkStyle();

    /**
     * Corresponding WML {@link NumberFormat}.
     *
     * @return number format
     */
    NumberFormat getNumberFormat();

    /**
     * Corresponding to <code>tplc</code> value.
     *
     * @return <code>tplc</code> value
     */
    String getId();

    /**
     * Display value of this list item.
     *
     * @param number current value
     * @return value of cuurent item
     */
    String getValue(int number);

    /**
     * Calculates and returns left indent value for given level;
     *
     * @param level current level
     * @return left indent value
     */
    long getLeftIndent(int level);

    /**
     * Calculates and returns hanging value for given level;
     *
     * @param level current level
     * @return hanging value
     */
    long getHangingValue(int level);

    /**
     * Returns tentative flag for current level.
     *
     * @param level current level
     * @return flag for tentative element or null
     */
    Boolean isTentative(int level);

    /**
     * Value of multi level type.
     *
     * @return multi level type
     */
    String getMultiLevelType();

    /**
     * Run properties if any.
     *
     * @return run properties
     */
    RPr getRPr();

    /**
     * Name of this ListItem.
     *
     * @return name of this list item.
     */
    String getName();
}
