package com.alphasystem.openxml.builder.wml;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.CTOverride;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.math.BigInteger;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.HeadingList.*;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadNumbering;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.loadStyles;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getCTRelBuilder;
import static java.lang.String.format;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.docx4j.openpackaging.contenttype.ContentTypes.WORDPROCESSINGML_DOCUMENT;
import static org.docx4j.openpackaging.contenttype.ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED;

/**
 * Fluent API for creating WordML package.
 *
 * @author sali
 */
public class WmlPackageBuilder {

    private final Logger logger = LoggerFactory.getLogger(getClass());
    private NumberingHelper numberingHelper;
    private WordprocessingMLPackage wmlPackage;
    private Numbering numbering;

    public WmlPackageBuilder() throws Docx4JException {
        this("META-INF/default.dotx");
    }

    public WmlPackageBuilder(boolean loadDefaultStyles) throws Docx4JException {
        wmlPackage = WordprocessingMLPackage.createPackage();
        if (loadDefaultStyles) {
            wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().setJaxbElement(loadStyles(null, "styles.xml"));
        }
        numberingHelper = new NumberingHelper();
        numberingHelper.populateDefaultNumbering();
        numbering = numberingHelper.getNumbering();
    }

    public WmlPackageBuilder(String templatePath) throws Docx4JException {
        URL url;
        final List<URL> urls;
        try {
            urls = WmlAdapter.readResources(templatePath);
            url = urls.get(0);
        } catch (IOException e) {
            throw new UncheckedIOException(e.getMessage(), e);
        }
        if (url == null) {
            throw new RuntimeException(format("Unable to open template \"%s\".", templatePath));
        }
        try {
            loadTemplate(templatePath, url);
        } catch (IOException e) {
            throw new UncheckedIOException(e.getMessage(), e);
        } catch (URISyntaxException e) {
            throw new RuntimeException(e.getMessage(), e);
        }
        numberingHelper = new NumberingHelper();
        numberingHelper.populateDefaultNumbering();
        numbering = numberingHelper.getNumbering();
    }

    private void loadTemplate(String templatePath, URL url) throws Docx4JException, IOException, URISyntaxException {
        try (InputStream is = url.openStream()) {
            wmlPackage = WordprocessingMLPackage.load(is);

            // Replace dotx content type with docx
            ContentTypeManager ctm = wmlPackage.getContentTypeManager();

            // Get <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"/>
            CTOverride override = ctm.getOverrideContentType().get(new URI("/word/document.xml")); // note this assumption

            String contentType = templatePath.endsWith("dotm") /* macro enabled? */ ? WORDPROCESSINGML_DOCUMENT_MACROENABLED : WORDPROCESSINGML_DOCUMENT;
            override.setContentType(contentType);

            // Create settings part, and init content
            DocumentSettingsPart dsp = new DocumentSettingsPart();
            CTSettings settings = Context.getWmlObjectFactory().createCTSettings();
            dsp.setJaxbElement(settings);
            wmlPackage.getMainDocumentPart().addTargetPart(dsp);

            // Create external rel
            RelationshipsPart rp = RelationshipsPart.createRelationshipsPartForPart(dsp);
            org.docx4j.relationships.Relationship rel = new org.docx4j.relationships.ObjectFactory().createRelationship();
            rel.setType(Namespaces.ATTACHED_TEMPLATE);
            rel.setTarget(url.toExternalForm());
            rel.setTargetMode("External");
            rp.addRelationship(rel); // addRelationship sets the rel's @Id

            settings.setAttachedTemplate(getCTRelBuilder().withId(rel.getId()).getObject());
        }
    }

    @SafeVarargs
    public final <T extends ListItem<T>> WmlPackageBuilder multiLevelHeading(T... items) {
        final int numberId = numberingHelper.populate(items);
        final BigInteger requiredValue = BigInteger.valueOf(numberId);
        T firstItem = items[0];
        for (T currentItem : items) {
            final String styleName = currentItem.getStyleName();
            if (styleName == null) {
                logger.error("No name defined in multi level heading item \"{}\"", firstItem.getName());
                continue;
            }
            final Style style = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().getStyleById(styleName);
            if (style == null) {
                logger.error("######################################################################################################################");
                logger.error("##### No style with name \"{}\" for \"{}\", possible reason is that style might not initialized first #####", styleName, currentItem.getName());
                logger.error("######################################################################################################################");
            } else {
                final PPrBase.NumPr.NumId numId = style.getPPr().getNumPr().getNumId();
                final BigInteger numIdVal = numId.getVal();
                if (!numIdVal.equals(requiredValue)) {
                    logger.info("Found number ID value of \"{}\" but requires value of \"{}\" for style \"{}\", changing it now.", numIdVal, numberId, styleName);
                    numId.setVal(requiredValue);
                }
            }
        } // end of for
        numbering = numberingHelper.getNumbering();
        return this;
    }

    public final <T extends HeadingList> WmlPackageBuilder multiLevelHeading(T... items) {
        final int numberId = numberingHelper.populate(items);
        for (int i = 0; i < items.length; i++) {
            final T currentItem = items[i];
            final String styleName = currentItem.getStyleName();
            final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
            Style style = styleDefinitionsPart.getStyleById(styleName);
            if (style == null) {
                throw new RuntimeException(format("No style found with id \"%s\"", styleName));
            }
            StyleBuilder styleBuilder = new StyleBuilder(style);
            PPrBuilder pPrBuilder = new PPrBuilder(styleBuilder.getObject().getPPr());
            Long level = i <= 0 ? null : Long.valueOf(i);
            final PPrBase.NumPr numPr = pPrBuilder.getNumPrBuilder().withNumId(Long.valueOf(numberId))
                    .withIlvl(level).getObject();
            pPrBuilder.withNumPr(numPr);
        }
        return this;
    }

    public WmlPackageBuilder multiLevelHeading() {
        // copy Heading1, TOCHeading is based on it
        final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
        final Style tocHeading = styleDefinitionsPart.getStyleById("TOCHeading");
        if (tocHeading != null) {
            String styleName = HEADING1.getStyleName();
            Style style = styleDefinitionsPart.getStyleById(styleName);
            if (style == null) {
                throw new RuntimeException(format("No style found with id \"%s\"", styleName));
            }
            StyleBuilder styleBuilder = new StyleBuilder(style, null);
            styleName = format("_%s", styleName);
            styleBuilder.withStyleId(styleName).withName(format("_", HEADING1.getName()))
                    .withUnhideWhenUsed(null).withSemiHidden(null).withHidden(BOOLEAN_DEFAULT_TRUE_TRUE);
            try {
                styleDefinitionsPart.getContents().getStyle().add(styleBuilder.getObject());
                new StyleBuilder(tocHeading).withBasedOn(styleName);
            } catch (Docx4JException e) {
                logger.warn("Unable to add style \"{}\" into style gallery.", styleName, e);
            }
        }
        return multiLevelHeading(HEADING1, HEADING2, HEADING3, HEADING4, HEADING5);
    }

    public WmlPackageBuilder styles(String... paths) {
        try {
            final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
            Styles styles = loadStyles(styleDefinitionsPart.getContents(), paths);
            styleDefinitionsPart.setJaxbElement(styles);
        } catch (Docx4JException e) {
            e.printStackTrace();
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
        mainDocumentPart.getContent().clear();
        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ndp.setJaxbElement(numbering);
        mainDocumentPart.addTargetPart(ndp);

        return wmlPackage;
    }
}
