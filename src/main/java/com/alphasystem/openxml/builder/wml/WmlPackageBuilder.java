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
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getCTRelBuilder;
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
    private Styles styles;
    private Numbering numbering;

    public WmlPackageBuilder() throws Docx4JException {
        this(true);
    }

    public WmlPackageBuilder(boolean loadDefaultStyles) throws Docx4JException {
        wmlPackage = WordprocessingMLPackage.createPackage();
        if (loadDefaultStyles) {
            styles = loadStyles(styles, "styles.xml");
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
            throw new RuntimeException(String.format("Unable to open template \"%s\".", templatePath));
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

            styles = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().getContents();
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
                    }
                    break;
                }
            } // end of for loop
            if (!styleFound) {
                logger.error("######################################################################################################################");
                logger.error("##### No style with name \"{}\" for \"{}\", possible reason is that style might not initialized first #####", styleName, currentItem.getName());
                logger.error("######################################################################################################################");
            }
        } // end of for
        numbering = numberingHelper.getNumbering();
        return this;
    }

    public WmlPackageBuilder multiLevelHeading() {
        styles = loadStyles(styles, "multi-level-heading/styles.xml");
        return multiLevelHeading(HEADING1, HEADING2, HEADING3, HEADING4, HEADING5);
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
