package com.alphasystem.openxml.builder.wml;

import com.alphasystem.commons.SystemException;
import com.alphasystem.commons.util.AppUtil;
import com.alphasystem.docx4j.builder.wml.PPrBuilder;
import com.alphasystem.docx4j.builder.wml.StyleBuilder;
import jakarta.xml.bind.JAXBException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.docx4j.Docx4jProperties;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.contenttype.CTOverride;
import org.docx4j.openpackaging.contenttype.ContentTypeManager;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.DocumentSettingsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.util.Objects;

import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getCTRelBuilder;
import static java.lang.String.format;
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
    private final NumberingHelper numberingHelper = NumberingHelper.getInstance();
    ;
    private WordprocessingMLPackage wmlPackage;

    public WmlPackageBuilder() throws SystemException {
        this(new WmlPackageInputs());
    }

    public WmlPackageBuilder(WmlPackageInputs inputs) throws SystemException {
        initPackage(inputs);
    }

    private void initPackage(WmlPackageInputs inputs) throws SystemException {
        final var pageSizePaper = inputs.getPageSizePaper();
        final var landscape = inputs.isLandscape();
        final var templatePath = inputs.getTemplatePath();
        final var loadDefaultStyles = inputs.isLoadDefaultStyles();
        if (Objects.nonNull(pageSizePaper)) {
            try {
                wmlPackage = WordprocessingMLPackage.createPackage(pageSizePaper, landscape);
            } catch (InvalidFormatException e) {
                logger.warn("Unable to load package from page zie: {}", pageSizePaper.value());
            }
        }
        if (Objects.nonNull(templatePath)) {
            try {
                wmlPackage = new TemplateLoader(templatePath).getPackage();
            } catch (SystemException e) {
                logger.warn("Unable to load package from template: {}", templatePath);
            }
        }
        if (Objects.isNull(wmlPackage)) {
            try {
                wmlPackage = WordprocessingMLPackage.createPackage();
            } catch (InvalidFormatException e) {
                throw new SystemException("Unable to create default package", e);
            }
        }
        if (loadDefaultStyles) {
            try {
                Styles styles = new StyleLoader(null, "META-INF/styles.xml").getStyles();
                wmlPackage.getMainDocumentPart().getStyleDefinitionsPart().setJaxbElement(styles);
            } catch (Exception ex) {
                logger.warn("Unable to load default styles", ex);
            }
        }
    }

    @SafeVarargs
    public final <T extends HeadingList<T>> WmlPackageBuilder multiLevelHeading(T... items) {
        final var numberId = numberingHelper.populate(items);
        for (int i = 0; i < items.length; i++) {
            final T currentItem = items[i];
            final String styleName = currentItem.getStyleName();
            final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
            Style style = styleDefinitionsPart.getStyleById(styleName);
            if (style == null) {
                throw new RuntimeException(format("No style found with id \"%s\"", styleName));
            }
            final var styleBuilder = new StyleBuilder(style);
            final PPr pPr = styleBuilder.getObject().getPPr();
            boolean setPPr = pPr == null;
            final var pPrBuilder = new PPrBuilder(pPr);
            Long level = i <= 0 ? null : (long) i;
            final PPrBase.NumPr numPr = pPrBuilder.getNumPrBuilder().withNumId(numberId)
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
                    .withUnhideWhenUsed((BooleanDefaultTrue) null).withSemiHidden((BooleanDefaultTrue) null)
                    .withHidden(BOOLEAN_DEFAULT_TRUE_TRUE);
            try {
                styleDefinitionsPart.getContents().getStyle().add(styleBuilder.getObject());
                new StyleBuilder(tocHeading).withBasedOn(styleName);
            } catch (Docx4JException e) {
                logger.warn("Unable to add style \"{}\" into style gallery.", styleName, e);
            }
        }
        return multiLevelHeading(Headings.HEADING1, Headings.HEADING2, Headings.HEADING3, Headings.HEADING4, Headings.HEADING5);
    }

    public WmlPackageBuilder styles(String... paths) throws SystemException {
        final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
        final Styles styles;
        try {
            styles = new StyleLoader(styleDefinitionsPart.getContents(), paths).getStyles();
        } catch (Docx4JException e) {
            throw new SystemException(e.getMessage(), e);
        }
        styleDefinitionsPart.setJaxbElement(styles);
        return this;
    }

    public WmlPackageBuilder styles(Styles... styles) throws SystemException {
        if (ArrayUtils.isEmpty(styles)) {
            return this;
        }
        final StyleDefinitionsPart styleDefinitionsPart = wmlPackage.getMainDocumentPart().getStyleDefinitionsPart();
        try {
            for (Styles style : styles) {
                styleDefinitionsPart.getContents().getStyle().addAll(style.getStyle());
            }
        } catch (Docx4JException e) {
            throw new SystemException(e);
        }
        return this;
    }

    public WordprocessingMLPackage getPackage() throws SystemException {
        final var mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.getContent().clear();
        try {
            final var ndp = new NumberingDefinitionsPart();
            ndp.setJaxbElement(numberingHelper.getNumbering());
            mainDocumentPart.addTargetPart(ndp);
        } catch (InvalidFormatException e) {
            throw new SystemException(e);
        }
        return wmlPackage;
    }

    private static class TemplateLoader {

        private WordprocessingMLPackage wmlPackage;
        private final String contentType;

        TemplateLoader(String templatePath) throws SystemException {
            templatePath = StringUtils.isBlank(templatePath) ? DEFAULT_TEMPLATE_PATH : templatePath;
            contentType = templatePath.endsWith("dotm") /* macro enabled? */ ? WORDPROCESSINGML_DOCUMENT_MACROENABLED
                    : WORDPROCESSINGML_DOCUMENT;
            loadTemplate(templatePath);
        }

        public WordprocessingMLPackage getPackage() {
            return wmlPackage;
        }

        private void loadTemplate(String templatePath) throws SystemException {
            final var wmlPackages = AppUtil.processResource(templatePath, this::loadTemplate);
            if (wmlPackages.isEmpty()) {
                return;
            }
            wmlPackage = wmlPackages.get(0);
        }

        private WordprocessingMLPackage loadTemplate(URL url) {
            try (final var is = url.openStream()) {
                final var wmlPackage = WordprocessingMLPackage.load(is);
                // Replace dotx content type with docx
                ContentTypeManager ctm = wmlPackage.getContentTypeManager();

                // Get <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"/>
                CTOverride override = ctm.getOverrideContentType().get(new URI("/word/document.xml")); // note this assumption

                override.setContentType(contentType);

                // Create settings part, and init content
                final var dsp = new DocumentSettingsPart();
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
                return wmlPackage;
            } catch (IOException | Docx4JException | URISyntaxException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static class StyleLoader {

        private Styles styles;

        StyleLoader(Styles styles, String... paths) throws SystemException {
            this.styles = styles;
            loadStyles(paths);
        }

        public Styles getStyles() {
            return styles;
        }

        private void loadStyles(String... paths) throws SystemException {
            if (Objects.isNull(paths)) {
                return;
            }
            for (String path : paths) {
                final var newStyles = AppUtil.processResource(path, StyleLoader::unmarshal);
                newStyles.forEach(otherStyles -> {
                    if (Objects.isNull(styles)) {
                        styles = otherStyles;
                    } else {
                        styles.getStyle().addAll(otherStyles.getStyle());
                    }
                });
            }
        }

        private static Styles unmarshal(URL url) {
            try (final var is = url.openStream()) {
                return (Styles) XmlUtils.unmarshal(is);
            } catch (IOException | JAXBException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public static class WmlPackageInputs {

        private String templatePath;
        private  boolean loadDefaultStyles;
        private boolean landscape;
        private PageSizePaper pageSizePaper;

        public WmlPackageInputs() {
        }

        public String getTemplatePath() {
            return templatePath;
        }

        public boolean isLoadDefaultStyles() {
            return loadDefaultStyles;
        }

        public boolean isLandscape() {
            return landscape;
        }

        public PageSizePaper getPageSizePaper() {
            return pageSizePaper;
        }

        public WmlPackageInputs withTemplatePath(String templatePath) {
            this.templatePath = templatePath;
            return this;
        }

        public WmlPackageInputs useDefaultTemplate() {
            this.templatePath = DEFAULT_TEMPLATE_PATH;
            return this;
        }

        public WmlPackageInputs useDefaultStyles() {
            this.loadDefaultStyles = true;
            return this;
        }

        public WmlPackageInputs useLandscape() {
            this.templatePath = DEFAULT_LANDSCAPE_TEMPLATE;
            this.landscape = true;
            return this;
        }

        public WmlPackageInputs withPageSizePaper(PageSizePaper pageSizePaper) {
            this.pageSizePaper = pageSizePaper;
            return this;
        }

        public WmlPackageInputs withPageSizePaper(String value) {
            final var paperSize = Docx4jProperties.getProperties().getProperty("docx4j.PageSize", value);
            try {
                pageSizePaper = PageSizePaper.valueOf(paperSize);
            } catch (IllegalArgumentException ex) {
                System.err.printf("Invalid page size value: %s%n", value);
            }
            return this;
        }
    }
}
