package com.alphasystem.openxml.gradleplugin;

import org.docx4j.wml.*;

import java.io.File;
import java.io.IOException;

public class FluentApiGeneratorMain {
    public static void main(String[] args) {
        FluentApiGenerator apiGenerator = new FluentApiGenerator("com.alphasystem.openxml.builder",
                "wml", "WmlBuilderFactory",
                P.class, P.Hyperlink.class, Tbl.class, Tr.class, Tc.class, R.class, Text.class, CTTabStop.class,
                Br.class, FldChar.class, SectPr.class, TblGridCol.class, CTBookmarkRange.class, CTBookmark.class,
                BooleanDefaultFalse.class, Styles.class, Style.class, Numbering.class, SdtBlock.class, CTSdtRow.class);


        try {
            apiGenerator.generate("org.docx4j.wml", new File("test"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
