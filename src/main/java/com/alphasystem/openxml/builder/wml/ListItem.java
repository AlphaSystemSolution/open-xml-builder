package com.alphasystem.openxml.builder.wml;

import org.docx4j.wml.NumberFormat;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRFontsBuilder;
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

    /**
     * ID of the list.
     *
     * @return list id
     */
    int getNumberId();

    /**
     * Name of the style.
     *
     * @return style name
     */
    String getStyleName();

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
     * Run properties if any.
     *
     * @return run properties
     */
    RPr getRPr();

    /**
     * Next list item. This value is used to create numbering levels for the numbering system.
     *
     * @return next item value
     */
    T getNext();

    static RFonts getRFonts(String ascii, String hAnsi) {
        return getRFonts(ascii, hAnsi, null);
    }

    static RFonts getRFonts(String ascii, String hAnsi, String cs) {
        return getRFontsBuilder().withAscii(ascii).withHAnsi(hAnsi).withCs(cs).withHint(DEFAULT).getObject();
    }
}
