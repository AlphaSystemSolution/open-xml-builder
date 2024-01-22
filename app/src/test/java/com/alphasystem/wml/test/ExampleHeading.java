package com.alphasystem.wml.test;

import com.alphasystem.commons.util.IdGenerator;
import com.alphasystem.openxml.builder.wml.HeadingList;

public class ExampleHeading extends HeadingList<ExampleHeading> {

    public ExampleHeading() {
        super("ExampleTitle", IdGenerator.nextId());
    }
}
