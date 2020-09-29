package com.alphasystem.openxml.gradleplugin;

import org.docx4j.wml.*;

import java.io.File;
import java.io.IOException;

public final class CodeGenerator {

    public static void wmlGenerator(File destPath) {
        FluentApiGenerator apiGenerator = new FluentApiGenerator("com.alphasystem.openxml.builder",
                "wml", "WmlBuilderFactory", P.class, P.Hyperlink.class, Tbl.class,
                Tr.class, Tc.class, R.class, R.Tab.class, Text.class, CTTabStop.class, Br.class, FldChar.class,
                SectPr.class, TblGridCol.class, CTBookmarkRange.class, CTBookmark.class, BooleanDefaultFalse.class,
                Styles.class, Style.class, Numbering.class, SdtBlock.class, CTSdtRow.class, CTSdtDocPart.class,
                CTHeight.class);

        try {
            apiGenerator.generate("org.docx4j.wml", destPath);
        } catch (IOException ex) {
            throw new RuntimeException("Unable to generate ", ex);
        }
    }
}
