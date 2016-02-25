package com.alphasystem.openxml.builder.wml;

import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.Styles;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadNumbering;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadStyles;

/**
 * Fluent API for creating WordML package.
 *
 * @author sali
 */
public class WmlPackageBuilder {

    private WordprocessingMLPackage wmlPackage;
    private Styles styles;
    private Numbering numbering;
    private boolean multiLevelHeading;

    public WmlPackageBuilder() throws InvalidFormatException {
        wmlPackage = WordprocessingMLPackage.createPackage();
        styles = loadStyles(styles, "styles.xml");
        numbering = loadNumbering(numbering, "numbering.xml");
    }

    public WmlPackageBuilder multiLevelHeading() {
        styles = loadStyles(styles, "multi-level-heading/styles.xml");
        numbering = loadNumbering(numbering, "multi-level-heading/numbering.xml");
        multiLevelHeading = true;
        return this;
    }

    public WmlPackageBuilder styles(String... paths) {
        styles = loadStyles(styles, paths);
        return this;
    }

    public WmlPackageBuilder styles(Styles customStyles) {
        styles.getStyle().addAll(customStyles.getStyle());
        return this;
    }

    public WmlPackageBuilder numbering(String... customNumberings) {
        numbering = loadNumbering(numbering, customNumberings);
        return this;
    }

    public WmlPackageBuilder numbering(Numbering customNumbering) {
        numbering.getAbstractNum().addAll(customNumbering.getAbstractNum());
        numbering.getNum().addAll(customNumbering.getNum());
        return this;
    }

    public WordprocessingMLPackage getPackage() throws InvalidFormatException {
        if (!multiLevelHeading) {
            // if multi level heading wasn't asked then load default headings
            styles = loadStyles(styles, "heading/styles.xml");
        }
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        StyleDefinitionsPart sdp = mainDocumentPart.getStyleDefinitionsPart();
        sdp.setJaxbElement(styles);

        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ndp.setJaxbElement(numbering);
        mainDocumentPart.addTargetPart(ndp);

        return wmlPackage;
    }
}
