package com.alphasystem.openxml.gradleplugin;

import org.docx4j.wml.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;

public class FluentApiGeneratorMain {
    private static final Logger LOGGER = LoggerFactory.getLogger(FluentApiGeneratorMain.class);

    static void main() {
        FluentApiGenerator apiGenerator = new FluentApiGenerator("com.alphasystem.openxml.builder",
                "wml", "WmlBuilderFactory");
        try {
            apiGenerator.generate("org.docx4j.wml", new File("test"),
                    P.class, P.Hyperlink.class, Tbl.class, Tr.class, Tc.class, R.class, Text.class, CTTabStop.class,
                    Br.class, FldChar.class, SectPr.class, TblGridCol.class, CTBookmarkRange.class, CTBookmark.class,
                    BooleanDefaultFalse.class, Styles.class, Style.class, Numbering.class, SdtBlock.class, CTSdtRow.class);
        } catch (IOException e) {
            LOGGER.error("Error occurred while generating fluent API", e);
        }
    }

}
