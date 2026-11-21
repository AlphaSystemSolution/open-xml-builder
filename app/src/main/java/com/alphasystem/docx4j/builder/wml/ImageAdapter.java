package com.alphasystem.docx4j.builder.wml;

import org.apache.commons.io.FilenameUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.P;

import java.net.URL;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;

import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;

public class ImageAdapter {

    private static final AtomicLong ID1 = new AtomicLong();
    private static final AtomicInteger ID2 = new AtomicInteger();

    private final WordprocessingMLPackage wordMLPackage;
    private final URL url;
    private String filenameHint;
    private String altText;

    public ImageAdapter(final WordprocessingMLPackage wordMLPackage, final URL url) {
        if (wordMLPackage == null || url == null) {
            throw new IllegalArgumentException("wordMLPackage and/or url cannot be null");
        }
        this.wordMLPackage = wordMLPackage;
        this.url = url;
    }

    public ImageAdapter withFilenameHint(String filenameHint) {
        this.filenameHint = filenameHint;
        return this;
    }

    public ImageAdapter withAltText(String altText) {
        this.altText = altText;
        return this;
    }

    public P insertImage() throws Exception {
        if (filenameHint == null || filenameHint.trim().isBlank()) {
            filenameHint = FilenameUtils.getName(url.getPath());
        }
        if (altText == null || altText.trim().isBlank()) {
            altText = FilenameUtils.removeExtension(filenameHint);
        }
        final BinaryPartAbstractImage imagePart;
        try (final var input = url.openStream()) {
            imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, input.readAllBytes());
        }
        final var imageInline = imagePart.createImageInline(filenameHint, altText, ID1.incrementAndGet(),
                ID2.incrementAndGet(), false);
        final var drawing = getDrawingBuilder().addAnchorOrInline(imageInline).getObject();
        return getPBuilder().addContent(getRBuilder().addContent(drawing).getObject()).getObject();
    }
}
