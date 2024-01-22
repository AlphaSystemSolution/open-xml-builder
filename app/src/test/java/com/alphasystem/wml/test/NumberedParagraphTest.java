package com.alphasystem.wml.test;

import com.alphasystem.commons.SystemException;
import com.alphasystem.openxml.builder.wml.*;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;
import org.testng.annotations.Test;

import java.util.Arrays;

import static com.alphasystem.commons.util.IdGenerator.nextId;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.getRBuilder;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getEmptyPara;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static java.lang.String.format;

/**
 * @author sali
 */
public class NumberedParagraphTest extends CommonTest {

    private final NumberingHelper numberingHelper = NumberingHelper.getInstance();

    /**
     * Create something like:
     * <pre>
     * <w:p>
     * <w:pPr>
     * <w:numPr>
     * <w:ilvl w:val="0"/>
     * <w:numId w:val="1"/>
     * </w:numPr>
     * </w:pPr>
     * <w:r>
     * <w:t>B</w:t>
     * </w:r>
     * </w:p>
     * </pre>
     *
     * @return p
     */
    private static P createNumberedParagraph(ListItem<?> listItem, long level, String paragraphText, boolean applyNumbering) {
        final var r = getRBuilder().addContent(getText(paragraphText)).getObject();
        final var pPr = WmlAdapter.getListParagraphProperties(listItem, level, applyNumbering);
        return getPBuilder().withPPr(pPr).addContent(r).getObject();
    }

    @Override
    protected String getFileName() {
        return "numbered-paragraph.docx";
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws SystemException {
        final var inputs = new WmlPackageBuilder.WmlPackageInputs().withTemplatePath("META-INF/Custom.dotx");
        return new WmlPackageBuilder(inputs)
                .styles("META-INF/example-caption/styles.xml")
                .multiLevelHeading(new ExampleHeading()).getPackage();
    }

    @Test
    public void createListParagraphs() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();

        long level = 0;
        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading1", "Numbered Lists"));
        mainDocumentPart.addObject(getEmptyPara());
        Arrays.stream(OrderedList.values()).forEach(listItem -> addList(mainDocumentPart, listItem, level));

        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading1", "Unordered Lists"));
        mainDocumentPart.addObject(getEmptyPara());
        Arrays.stream(UnorderedList.values()).forEach(listItem -> addList(mainDocumentPart, listItem, level));

        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading1", "Multi-Level lists"));
        mainDocumentPart.addObject(getEmptyPara());
        Arrays.stream(OrderedList.values()).forEach(listItem -> addMultiLevelList(mainDocumentPart, listItem));
    }

    @Test(dependsOnMethods = "createListParagraphs")
    public void createListByStyleName() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading1", "Creating list by style name"));
        mainDocumentPart.addObject(getEmptyPara());

        String styleName = "upperroman";
        var listItem = numberingHelper.getListItem(styleName);
        listItem.setNumberId(restartNumbering(mainDocumentPart, listItem));
        addList(mainDocumentPart, listItem, 0L);

        mainDocumentPart.addObject(getEmptyPara());
        styleName = "diamond";
        listItem = numberingHelper.getListItem(styleName);
        listItem.setNumberId(restartNumbering(mainDocumentPart, listItem));
        addList(mainDocumentPart, listItem, 0L);
    }

    @Test(dependsOnMethods = "createListByStyleName")
    public void exampleTitle() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Example");
        mainDocumentPart.addObject(WmlAdapter.getParagraph("Some text"));
        mainDocumentPart.addObject(getEmptyPara());
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Example 2");
        mainDocumentPart.addObject(WmlAdapter.getParagraph("Some more text"));
        mainDocumentPart.addObject(getEmptyPara());
    }

    private void addList(final MainDocumentPart mainDocumentPart, ListItem<?> listItem, long level) {
        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading2", format("List of type: %s", listItem.getStyleName())));
        mainDocumentPart.addObject(getEmptyPara());
        mainDocumentPart.addObject(createNumberedParagraph(listItem, level, nextId(), true));
        mainDocumentPart.addObject(createNumberedParagraph(listItem, level, nextId(), false));
        mainDocumentPart.addObject(createNumberedParagraph(listItem, level, nextId(), true));
        mainDocumentPart.addObject(getEmptyPara());
    }

    private void addMultiLevelList(final MainDocumentPart mainDocumentPart, ListItem<?> listItem) {
        listItem.setNumberId(restartNumbering(mainDocumentPart, listItem));

        mainDocumentPart.addObject(WmlAdapter.getParagraphWithStyle("Heading2", format("List of type: %s", listItem.getStyleName())));
        for (int level = 0; level < 5; level++) {
            mainDocumentPart.addObject(createNumberedParagraph(listItem, level, nextId(), true));
        }
        mainDocumentPart.addObject(getEmptyPara());
    }

    private long restartNumbering(final MainDocumentPart mainDocumentPart, ListItem<?> listItem) {
        return mainDocumentPart.getNumberingDefinitionsPart().restart(listItem.getNumberId(), 0, 1);
    }
}
