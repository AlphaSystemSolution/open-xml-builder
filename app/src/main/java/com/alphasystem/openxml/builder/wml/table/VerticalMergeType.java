package com.alphasystem.openxml.builder.wml.table;

public enum VerticalMergeType {

    RESTART("restart"), CONTINUE(null), NONE(null);

    private final String value;

    VerticalMergeType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}