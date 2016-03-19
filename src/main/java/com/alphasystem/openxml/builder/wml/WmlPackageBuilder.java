package com.alphasystem.openxml.builder.wml;

import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.PPrBase;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigInteger;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadNumbering;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadStyles;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * Fluent API for creating WordML package.
 *
 * @author sali
 */
public class WmlPackageBuilder {

    private final Logger logger = LoggerFactory.getLogger(getClass());
    private NumberingHelper numberingHelper;
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
        numberingHelper = new NumberingHelper();
        numberingHelper.populateDefaultNumbering();
        numbering = numberingHelper.getNumbering();
    }

    public <T extends ListItem<T>> WmlPackageBuilder multiLevelHeading(T item) {
        final int numberId = numberingHelper.populate(item);
        final BigInteger requiredValue = BigInteger.valueOf(numberId);
        T currentItem = item;
        while (currentItem != null) {
            final String styleName = currentItem.getStyleName();
            if (styleName == null) {
                logger.error("No name defined in multi level heading item \"{}\"", item.getName());
                continue;
            }
            boolean styleFound = false;
            final List<Style> styleList = styles.getStyle();
            for (Style style : styleList) {
                if (style.getStyleId().equals(styleName)) {
                    styleFound = true;
                    final PPrBase.NumPr.NumId numId = style.getPPr().getNumPr().getNumId();
                    final BigInteger numIdVal = numId.getVal();
                    if (!numIdVal.equals(requiredValue)) {
                        logger.info("Found number ID value of \"{}\" but requires value of \"{}\" for style \"{}\", changing it now.", numIdVal, numberId, styleName);
                        numId.setVal(requiredValue);
                        break;
                    }
                }
            } // end of for loop
            if (!styleFound) {
                logger.error("######################################################################################################################");
                logger.error("##### No style with name \"{}\" for \"{}\", possible reason is that style might not initialized first #####", styleName, currentItem.getName());
                logger.error("######################################################################################################################");
            }
            currentItem = currentItem.getNext();
        } // end of while
        numbering = numberingHelper.getNumbering();
        return this;
    }

    public WmlPackageBuilder multiLevelHeading() {
        styles = loadStyles(styles, "multi-level-heading/styles.xml");
        return multiLevelHeading(HeadingList.HEADING1);
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
