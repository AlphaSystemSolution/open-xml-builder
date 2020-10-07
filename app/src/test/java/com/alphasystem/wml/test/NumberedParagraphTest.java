package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.*;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.P;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import java.io.File;

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
public class NumberedParagraphTest extends CommonTest {

    private static final String PARENT_PATH = "build";

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
    private static P createNumberedParagraph(long listId, long level, String paragraphText, boolean applyNumbering) {
        final RBuilder rBuilder = getRBuilder().addContent(getText(paragraphText));
        return getPBuilder().withPPr(WmlAdapter.getListParagraphProperties(listId, level, applyNumbering))
                .addContent(rBuilder.getObject()).getObject();
    }

    private WordprocessingMLPackage wmlPackage;

    @Override
    protected String getFileName() {
        return null;
    }

    @Override
    protected WordprocessingMLPackage loadWmlPackage() throws Docx4JException {
        return null;
    }

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = WmlPackageBuilder.createPackage().getPackage();
        } catch (Docx4JException e) {
            fail(e.getMessage(), e);
        }
    }

    @AfterClass
    public void tearDown() {
        try {
            final File file = get(PARENT_PATH, "test.docx").toFile();
            save(file, wmlPackage);
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
        OrderedList orderedListItem = OrderedList.getByStyleName(styleName);
        addList(mainDocumentPart, orderedListItem.getNumberId(), 0L);

        mainDocumentPart.addObject(getEmptyPara());
        styleName = "diamond";
        mainDocumentPart.addParagraphOfText(format("Creating list by style name \"%s\".", styleName));
        mainDocumentPart.addObject(getEmptyPara());
        UnorderedList unorderedListItem = UnorderedList.getByStyleName(styleName);
        addList(mainDocumentPart, unorderedListItem.getNumberId(), 0L);
    }

    @Test(dependsOnMethods = "createListByStyleName")
    public void exampleTitle() {
        final MainDocumentPart mainDocumentPart = wmlPackage.getMainDocumentPart();
        mainDocumentPart.addObject(getEmptyPara());
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Example");
        mainDocumentPart.addObject(getEmptyPara());
        mainDocumentPart.addStyledParagraphOfText("ExampleTitle", "Example 2");
    }

    private void addList(final MainDocumentPart mainDocumentPart, long numId, long ilvl) {
        mainDocumentPart.addParagraphOfText(format("List of type: %s", numId));
        mainDocumentPart.addObject(createNumberedParagraph(numId, ilvl, nextId(), true));
        mainDocumentPart.addObject(createNumberedParagraph(numId, ilvl, nextId(), false));
        mainDocumentPart.addObject(createNumberedParagraph(numId, ilvl, nextId(), true));
        mainDocumentPart.addObject(getEmptyPara());
    }

    private void addMultiLevelList(final MainDocumentPart mainDocumentPart, long numId, long actualNumId) {
        mainDocumentPart.addParagraphOfText(format("Multi-Level List of type: %s", actualNumId));
        for (int level = 0; level < 5; level++) {
            mainDocumentPart.addObject(createNumberedParagraph(numId, level, nextId(), true));
        }
        mainDocumentPart.addObject(getEmptyPara());
    }
}
