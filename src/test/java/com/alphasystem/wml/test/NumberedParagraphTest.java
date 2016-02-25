package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.OrderedListItem;
import com.alphasystem.openxml.builder.wml.RBuilder;
import com.alphasystem.openxml.builder.wml.UnorderedListItem;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.P;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.*;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRBuilder;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static java.nio.file.Paths.get;
import static org.testng.Assert.fail;

/**
 * @author sali
 */
public class NumberedParagraphTest {

    private static final String PARENT_PATH = "C:\\Users\\sali\\git-hub\\AlphaSystemSolution\\open-xml-builder\\build";

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
    private static P createNumberedParagraph(long listId, long number, String paragraphText) {
        final RBuilder rBuilder = getRBuilder().addContent(getText(paragraphText));
        return getPBuilder().withPPr(WmlAdapter.getListParagraphProperties(listId, number))
                .addContent(rBuilder.getObject()).getObject();
    }

    private WordprocessingMLPackage wmlPackage;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = WmlAdapter.createNewDoc(true);
        } catch (InvalidFormatException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            save(get(PARENT_PATH, "test.docx").toFile(), wmlPackage);
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    @Test
    public void createListParagraphs() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        final NumberingDefinitionsPart ndp = mainDocumentPart.getNumberingDefinitionsPart();

        long level = 0;
        int start = 1;
        int end = 5;
        mainDocumentPart.addParagraphOfText("Numbered Lists");
        mainDocumentPart.addObject(getEmptyPara());
        for (int i = start; i <= end; i++) {
            addList(mainDocumentPart, i, level);
        }

        start = 6;
        end = 10;
        mainDocumentPart.addParagraphOfText("Unordered Lists");
        mainDocumentPart.addObject(getEmptyPara());
        for (int i = start; i <= end; i++) {
            addList(mainDocumentPart, i, level);
        }

        start = 1;
        end = 10;
        for (int i = start; i <= end; i++) {
            addMultiLevelList(mainDocumentPart, ndp.restart(i, 0, 1), i);
        }
    }

    @Test(dependsOnMethods = "createListParagraphs")
    public void createListByStyleName() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        String styleName = "upperroman";
        mainDocumentPart.addParagraphOfText(format("Creating list by style name \"%s\".", styleName));
        mainDocumentPart.addObject(getEmptyPara());
        OrderedListItem orderedListItem = OrderedListItem.getByStyleName(styleName);
        addList(mainDocumentPart, orderedListItem.getNumberId(), 0L);

        mainDocumentPart.addObject(getEmptyPara());
        styleName = "diamond";
        mainDocumentPart.addParagraphOfText(format("Creating list by style name \"%s\".", styleName));
        mainDocumentPart.addObject(getEmptyPara());
        UnorderedListItem unorderedListItem = UnorderedListItem.getByStyleName(styleName);
        addList(mainDocumentPart, unorderedListItem.getNumberId(), 0L);
    }

    @Test(dependsOnMethods = "createListByStyleName")
    public void createMultiLevelHeading() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());

        for (int i = 1; i <= 5; i++) {
            String style = format("Heading%s", i);
            mainDocumentPart.addStyledParagraphOfText(style, style);
        }
    }

    private void addList(final MainDocumentPart mainDocumentPart, long numId, long ilvl) {
        mainDocumentPart.addParagraphOfText(format("List of type: %s", numId));
        mainDocumentPart.addObject(createNumberedParagraph(numId, ilvl, nextId()));
        mainDocumentPart.addObject(createNumberedParagraph(numId, ilvl, nextId()));
        mainDocumentPart.addObject(getEmptyPara());
    }

    private void addMultiLevelList(final MainDocumentPart mainDocumentPart, long numId, long actualNumId) {
        mainDocumentPart.addParagraphOfText(format("Multi-Level List of type: %s", actualNumId));
        for (int i = 0; i < 5; i++) {
            mainDocumentPart.addObject(createNumberedParagraph(numId, i, nextId()));
        }
        mainDocumentPart.addObject(getEmptyPara());
    }
}
