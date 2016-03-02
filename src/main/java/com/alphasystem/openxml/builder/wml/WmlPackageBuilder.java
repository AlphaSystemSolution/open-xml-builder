package com.alphasystem.openxml.builder.wml;

import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.Styles;

import static com.alphasystem.openxml.builder.wml.NumberingHelper.getDefaultNumbering;
import static com.alphasystem.openxml.builder.wml.NumberingHelper.getMultiLevelHeadingNumbering;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadNumbering;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadStyles;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * Fluent API for creating WordML package.
 *
 * @author sali
 */
public class WmlPackageBuilder {

    private WordprocessingMLPackage wmlPackage;
    private Styles styles;
    private Numbering numbering;

    public WmlPackageBuilder() throws InvalidFormatException {
        this(true);
    }

    public WmlPackageBuilder(boolean loadDefaultStyles) throws InvalidFormatException {
        wmlPackage = WordprocessingMLPackage.createPackage();
        if (loadDefaultStyles) {
            styles = loadStyles(styles, "styles.xml");
        }
        numbering = getDefaultNumbering();
    }

    public WmlPackageBuilder multiLevelHeading() {
        styles = loadStyles(styles, "multi-level-heading/styles.xml");
        return numbering(getMultiLevelHeadingNumbering());
    }

    public WmlPackageBuilder styles(String... paths) {
        styles = loadStyles(styles, paths);
        return this;
    }

    public WmlPackageBuilder styles(Styles... styles) {
        if (!isEmpty(styles)) {
            for (Styles style : styles) {
                this.styles.getStyle().addAll(style.getStyle());
            }
        }
        return this;
    }

    public WmlPackageBuilder numbering(String... customNumberings) {
        numbering = loadNumbering(numbering, customNumberings);
        return this;
    }

    public WmlPackageBuilder numbering(Numbering... numberings) {
        if (!isEmpty(numberings)) {
            for (Numbering numbering : numberings) {
                this.numbering.getAbstractNum().addAll(numbering.getAbstractNum());
                this.numbering.getNum().addAll(numbering.getNum());
            }
        }
        return this;
    }

    public WordprocessingMLPackage getPackage() throws InvalidFormatException {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        StyleDefinitionsPart sdp = mainDocumentPart.getStyleDefinitionsPart();
        sdp.setJaxbElement(styles);

        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ndp.setJaxbElement(numbering);
        mainDocumentPart.addTargetPart(ndp);

        return wmlPackage;
    }
}
