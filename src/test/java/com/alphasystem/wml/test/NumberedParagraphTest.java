package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.*;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.P;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

import javax.xml.bind.JAXBException;

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

    private static final String PARENT_PATH = "C:\\Users\\sali\\git-hub\\AlphaSystemSolution\\open-xml-builder\\target";

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
    private static P createNumberedParagraph(long listId, long number, String paragraphText, boolean applyNumbering) {
        final RBuilder rBuilder = getRBuilder().addContent(getText(paragraphText));
        return getPBuilder().withPPr(WmlAdapter.getListParagraphProperties(listId, number, applyNumbering))
                .addContent(rBuilder.getObject()).getObject();
    }

    private WordprocessingMLPackage wmlPackage;

    @BeforeClass
    public void setup() {
        try {
            wmlPackage = new WmlPackageBuilder().getPackage();
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
        OrderedList orderedListItem = OrderedList.getByStyleName(styleName);
        addList(mainDocumentPart, orderedListItem.getNumberId(), 0L);

        mainDocumentPart.addObject(getEmptyPara());
        styleName = "diamond";
        mainDocumentPart.addParagraphOfText(format("Creating list by style name \"%s\".", styleName));
        mainDocumentPart.addObject(getEmptyPara());
        UnorderedList unorderedListItem = UnorderedList.getByStyleName(styleName);
        addList(mainDocumentPart, unorderedListItem.getNumberId(), 0L);
    }

    @Test
    public void testUnMarshall() {
        String s = "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
                " xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" " +
                "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
                "xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">\n" +
                "    <w:sdtPr>\n" +
                "        <w:id w:val=\"770900403\"/>\n" +
                "        <w:docPartObj>\n" +
                "            <w:docPartGallery w:val=\"Table of Contents\"/>\n" +
                "            <w:docPartUnique/>\n" +
                "        </w:docPartObj>\n" +
                "    </w:sdtPr>\n" +
                "    <w:sdtEndPr>\n" +
                "        <w:rPr>\n" +
                "            <w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\"\n" +
                "                      w:cstheme=\"minorBidi\"/>\n" +
                "            <w:b/>\n" +
                "            <w:bCs/>\n" +
                "            <w:noProof/>\n" +
                "            <w:color w:val=\"auto\"/>\n" +
                "            <w:sz w:val=\"22\"/>\n" +
                "            <w:szCs w:val=\"22\"/>\n" +
                "        </w:rPr>\n" +
                "    </w:sdtEndPr>\n" +
                "    <w:sdtContent>\n" +
                "        <w:p w:rsidR=\"0015683B\" w:rsidRDefault=\"0015683B\">\n" +
                "            <w:pPr>\n" +
                "                <w:pStyle w:val=\"TOCHeading\"/>\n" +
                "            </w:pPr>\n" +
                "            <w:r>\n" +
                "                <w:t>Contents</w:t>\n" +
                "            </w:r>\n" +
                "        </w:p>\n" +
                "        <w:p w:rsidR=\"0015683B\" w:rsidRDefault=\"0015683B\">\n" +
                "            <w:pPr>\n" +
                "                <w:pStyle w:val=\"TOC1\"/>\n" +
                "                <w:tabs>\n" +
                "                    <w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"9350\"/>\n" +
                "                </w:tabs>\n" +
                "                <w:rPr>\n" +
                "                    <w:noProof/>\n" +
                "                </w:rPr>\n" +
                "            </w:pPr>\n" +
                "            <w:r>\n" +
                "                <w:fldChar w:fldCharType=\"begin\"/>\n" +
                "            </w:r>\n" +
                "            <w:r>\n" +
                "                <w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText>\n" +
                "            </w:r>\n" +
                "            <w:r>\n" +
                "                <w:fldChar w:fldCharType=\"separate\"/>\n" +
                "            </w:r>\n" +
                "        </w:p>\n" +
                "        <w:p w:rsidR=\"0015683B\" w:rsidRDefault=\"0015683B\">\n" +
                "            <w:r>\n" +
                "                <w:rPr>\n" +
                "                    <w:b/>\n" +
                "                    <w:bCs/>\n" +
                "                    <w:noProof/>\n" +
                "                </w:rPr>\n" +
                "                <w:fldChar w:fldCharType=\"end\"/>\n" +
                "            </w:r>\n" +
                "        </w:p>\n" +
                "    </w:sdtContent>\n" +
                "</w:sdt>";
        try {
            final Object o = XmlUtils.unmarshalString(s);
            System.out.println("///////////////////////////////////// " + o);
        } catch (JAXBException e) {
            e.printStackTrace();
        }
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
        for (int i = 0; i < 5; i++) {
            mainDocumentPart.addObject(createNumberedParagraph(numId, i, nextId(), true));
        }
        mainDocumentPart.addObject(getEmptyPara());
    }
}
