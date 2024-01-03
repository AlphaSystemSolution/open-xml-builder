package com.alphasystem.wml.test;

import com.alphasystem.openxml.builder.wml.HeadingList;
import com.alphasystem.util.IdGenerator;

public class ExampleHeading extends HeadingList<ExampleHeading> {

    public ExampleHeading() {
        super("ExampleTitle", IdGenerator.nextId());
    }
}
