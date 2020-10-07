package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.WmlAdapter;
import org.testng.annotations.Test;

/**
 * @author sali
 */
public abstract class CustomStylesTest extends CommonTest {

    @Test
    public void testCustomStyle(){
        getMainDocumentPart().addObject(WmlAdapter.getParagraphWithStyle("ExampleTitle", "Example"));
    }
}
