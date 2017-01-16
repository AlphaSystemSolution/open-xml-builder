package com.alphasystem.openxml.builder.wml;

import org.apache.commons.lang3.ArrayUtils;
import org.docx4j.Docx4jProperties;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageSizePaper;
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
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getCTRelBuilder;
import static java.lang.String.format;
import static org.apache.commons.lang3.StringUtils.isBlank;
import static org.docx4j.openpackaging.contenttype.ContentTypes.WORDPROCESSINGML_DOCUMENT;
import static org.docx4j.openpackaging.contenttype.ContentTypes.WORDPROCESSINGML_DOCUMENT_MACROENABLED;

/**
 * Fluent API for creating WordML package.
 *
 * @author sali
 */
public class WmlPackageBuilder {

    private static final String DEFAULT_TEMPLATE_PATH = "META-INF/default.dotx";
    private static final String DEFAULT_LANDSCAPE_TEMPLATE = "META-INF/default-landscape.dotx";
    private final Logger logger = LoggerFactory.getLogger(getClass());
    private NumberingHelper numberingHelper;
    private WordprocessingMLPackage wmlPackage;

    public static WmlPackageBuilder createPackage() throws Docx4JException {
        return createPackage(null);
    }

    public static WmlPackageBuilder createPackage(boolean landscape) throws Docx4JException {
        return new WmlPackageBuilder(landscape, true);
    }

    public static WmlPackageBuilder createPackage(String templatePath) throws Docx4JException {
        return new WmlPackageBuilder(templatePath);
    }

    public WmlPackageBuilder(boolean loadDefaultStyles) throws Docx4JException {
        this(null, false, loadDefaultStyles);
    }

    public WmlPackageBuilder(PageSizePaper sz, boolean landscape, boolean loadDefaultStyles) throws Docx4JException {
        if (sz == null) {
            String paperSize = Docx4jProperties.getProperties().getProperty("docx4j.PageSize", "A4");
            try {
                sz = PageSizePaper.valueOf(paperSize);
            } catch (IllegalArgumentException e) {
                sz = PageSizePaper.A4;
            }
        }
        wmlPackage = WordprocessingMLPackage.createPackage(sz, landscape);
        if (loadDefaultStyles) {
            wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().setJaxbElement(WmlAdapter.loadStyles(null, "styles.xml"));
        }
        numberingHelper = new NumberingHelper();
        numberingHelper.populateDefaultNumbering();
    }

    private WmlPackageBuilder(String templatePath) throws Docx4JException {
        templatePath = isBlank(templatePath) ? DEFAULT_TEMPLATE_PATH : templatePath;
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
    }

    private WmlPackageBuilder(boolean landscape, @SuppressWarnings("unused") boolean dummy) throws Docx4JException {
        this(landscape ? DEFAULT_LANDSCAPE_TEMPLATE : DEFAULT_TEMPLATE_PATH);
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
    public final <T extends HeadingList<T>> WmlPackageBuilder multiLevelHeading(T... items) {
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
            final PPr pPr = styleBuilder.getObject().getPPr();
            boolean setPPr = pPr == null;
            PPrBuilder pPrBuilder = new PPrBuilder(pPr);
            Long level = i <= 0 ? null : (long) i;
            final PPrBase.NumPr numPr = pPrBuilder.getNumPrBuilder().withNumId(Long.valueOf(numberId))
                    .withIlvl(level).getObject();
            pPrBuilder.withNumPr(numPr);
            if (setPPr) {
                styleBuilder.withPPr(pPrBuilder.getObject());
            }
        }
        return this;
    }

    public WmlPackageBuilder multiLevelHeading() {
        // copy Heading1, TOCHeading is based on it
        final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
        final Style tocHeading = styleDefinitionsPart.getStyleById("TOCHeading");
        if (tocHeading != null) {
            String styleName = Headings.HEADING1.getStyleName();
            Style style = styleDefinitionsPart.getStyleById(styleName);
            if (style == null) {
                throw new RuntimeException(format("No style found with id \"%s\"", styleName));
            }
            StyleBuilder styleBuilder = new StyleBuilder(style, null);
            styleName = format("_%s", styleName);
            styleBuilder.withStyleId(styleName).withName(format("_%s", Headings.HEADING1.getName()))
                    .withUnhideWhenUsed((BooleanDefaultTrue) null).withSemiHidden((BooleanDefaultTrue) null).withHidden(BOOLEAN_DEFAULT_TRUE_TRUE);
            try {
                styleDefinitionsPart.getContents().getStyle().add(styleBuilder.getObject());
                new StyleBuilder(tocHeading).withBasedOn(styleName);
            } catch (Docx4JException e) {
                logger.warn("Unable to add style \"{}\" into style gallery.", styleName, e);
            }
        }
        return multiLevelHeading(Headings.HEADING1, Headings.HEADING2, Headings.HEADING3, Headings.HEADING4, Headings.HEADING5);
    }

    public WmlPackageBuilder styles(String... paths) {
        try {
            final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
            Styles styles = WmlAdapter.loadStyles(styleDefinitionsPart.getContents(), paths);
            styleDefinitionsPart.setJaxbElement(styles);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        return this;
    }

    public WmlPackageBuilder styles(Styles... styles) {
        if (ArrayUtils.isEmpty(styles)) {
            return this;
        }
        final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
        try {
            for (Styles style : styles) {
                styleDefinitionsPart.getContents().getStyle().addAll(style.getStyle());
            }
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
        return this;
    }

    public WordprocessingMLPackage getPackage() throws InvalidFormatException {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.getContent().clear();
        NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
        ndp.setJaxbElement(numberingHelper.getNumbering());
        mainDocumentPart.addTargetPart(ndp);

        return wmlPackage;
    }
}
