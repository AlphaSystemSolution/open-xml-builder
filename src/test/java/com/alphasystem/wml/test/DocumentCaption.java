package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.HeadingList;

import static java.lang.String.format;

/**
 * @author sali
 */
public abstract class DocumentCaption extends HeadingList<DocumentCaption> {

    public static final DocumentCaption EXAMPLE = new DocumentCaption("ExampleTitle") {

        @Override
        public String getValue(int i) {
            return format("Example %%%s.", i);
        }

        @Override
        public String getName() {
            return "EXAMPLE";
        }

    };

    DocumentCaption(String styleName) {
        super(styleName);
    }

    @Override
    public long getLeftIndent(int i) {
        return 432;
    }

}
