package com.alphasystem.docx4j.builder.wml;

import org.apache.batik.transcoder.TranscoderException;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.apache.commons.io.FilenameUtils;
import org.docx4j.dml.CTPoint2D;
import org.docx4j.dml.wordprocessingDrawing.Anchor;
import org.docx4j.dml.wordprocessingDrawing.CTPosH;
import org.docx4j.dml.wordprocessingDrawing.CTPosV;
import org.docx4j.dml.wordprocessingDrawing.CTWrapNone;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.dml.wordprocessingDrawing.STAlignH;
import org.docx4j.dml.wordprocessingDrawing.STAlignV;
import org.docx4j.dml.wordprocessingDrawing.STRelFromH;
import org.docx4j.dml.wordprocessingDrawing.STRelFromV;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.R;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.net.URL;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;

import static com.alphasystem.docx4j.builder.wml.WmlBuilderFactory.*;
import static org.docx4j.UnitsOfMeasurement.pxToTwipDouble;
import static org.docx4j.UnitsOfMeasurement.twipToEMU;

public class ImageAdapter {

    private static final AtomicLong ID1 = new AtomicLong();
    private static final AtomicInteger ID2 = new AtomicInteger();

    private final WordprocessingMLPackage wordMLPackage;
    private URL url;
    private String filenameHint;
    private String altText;
    private Double width;
    private Double height;
    private STRelFromH horizontalRelativeFrom;
    private STAlignH horizontalAlign;
    private STRelFromV verticalRelativeFrom;
    private STAlignV verticalAlign;

    public ImageAdapter(final WordprocessingMLPackage wordMLPackage) {
        if (wordMLPackage == null) {
            throw new IllegalArgumentException("wordMLPackage cannot be null");
        }
        this.wordMLPackage = wordMLPackage;
    }

    public ImageAdapter withUrl(URL url) {
        if (url == null) {
            throw new IllegalArgumentException("url cannot be null");
        }
        this.url = url;
        return this;
    }

    public ImageAdapter withFilenameHint(String filenameHint) {
        this.filenameHint = filenameHint;
        return this;
    }

    public ImageAdapter withAltText(String altText) {
        this.altText = altText;
        return this;
    }

    public ImageAdapter withWidth(Double width) {
        this.width = width;
        return this;
    }

    public ImageAdapter withHeight(Double height) {
        this.height = height;
        return this;
    }

    /**
     * Renders the image as a floating object aligned horizontally relative to the given reference.
     * When set (or {@link #withVerticalAlignment} is used), the drawing is emitted as a
     * {@code wp:anchor} instead of {@code wp:inline}.
     */
    public ImageAdapter withHorizontalAlignment(STRelFromH relativeFrom, STAlignH align) {
        this.horizontalRelativeFrom = relativeFrom;
        this.horizontalAlign = align;
        return this;
    }

    /**
     * Renders the image as a floating object aligned vertically relative to the given reference
     * (e.g. {@link STRelFromV#LINE} to align with the surrounding text line).
     */
    public ImageAdapter withVerticalAlignment(STRelFromV relativeFrom, STAlignV align) {
        this.verticalRelativeFrom = relativeFrom;
        this.verticalAlign = align;
        return this;
    }

    public R insertImage() throws Exception {
        if (url == null) {
            throw new IllegalArgumentException("URL cannot be null");
        }
        if (filenameHint == null || filenameHint.trim().isBlank()) {
            filenameHint = FilenameUtils.getName(url.getPath());
        }
        if (altText == null || altText.trim().isBlank()) {
            altText = FilenameUtils.removeExtension(filenameHint);
        }
        var imageBytes = readImageBytes();
        if (isSvg(imageBytes)) {
            imageBytes = convertSvgToPng(imageBytes);
            filenameHint = FilenameUtils.removeExtension(filenameHint) + ".png";
        }
        final var imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, imageBytes);

        // The (id1, id2, cx, link) overload expects cx in twips and derives height
        // from the aspect ratio; the (id1, id2, cx, cy, link) overload expects EMU.
        final Long widthTwips = width == null ? null : Math.round(pxToTwipDouble(width));
        Inline imageInline;
        if (widthTwips != null && height != null) {
            imageInline = imagePart.createImageInline(filenameHint, altText, ID1.incrementAndGet(),
                    ID2.incrementAndGet(), twipToEMU(widthTwips), twipToEMU(pxToTwipDouble(height)), false);
        } else if (widthTwips != null) {
            imageInline = imagePart.createImageInline(filenameHint, altText, ID1.incrementAndGet(),
                    ID2.incrementAndGet(), widthTwips, false);
        } else {
            imageInline = imagePart.createImageInline(filenameHint, altText, ID1.incrementAndGet(),
                    ID2.incrementAndGet(), false);
        }

        final var anchorOrInline = isAnchored() ? toAnchor(imageInline) : imageInline;
        reset();
        final var drawing = getDrawingBuilder().addAnchorOrInline(anchorOrInline).getObject();
        return getRBuilder().addContent(drawing).getObject();
    }

    private boolean isAnchored() {
        return horizontalAlign != null || verticalAlign != null;
    }

    private Anchor toAnchor(Inline inline) {
        final var positionH = new CTPosH();
        positionH.setRelativeFrom(horizontalRelativeFrom == null ? STRelFromH.COLUMN : horizontalRelativeFrom);
        if (horizontalAlign != null) {
            positionH.setAlign(horizontalAlign);
        } else {
            positionH.setPosOffset(0);
        }

        final var positionV = new CTPosV();
        positionV.setRelativeFrom(verticalRelativeFrom == null ? STRelFromV.LINE : verticalRelativeFrom);
        if (verticalAlign != null) {
            positionV.setAlign(verticalAlign);
        } else {
            positionV.setPosOffset(0);
        }

        final var simplePos = new CTPoint2D();
        simplePos.setX(0);
        simplePos.setY(0);

        final var anchor = new Anchor();
        anchor.setSimplePosAttr(false);
        anchor.setSimplePos(simplePos);
        anchor.setRelativeHeight(0);
        anchor.setBehindDoc(false);
        anchor.setLocked(false);
        anchor.setLayoutInCell(true);
        anchor.setAllowOverlap(true);
        anchor.setDistT(0L);
        anchor.setDistB(0L);
        anchor.setDistL(0L);
        anchor.setDistR(0L);
        anchor.setPositionH(positionH);
        anchor.setPositionV(positionV);
        anchor.setExtent(inline.getExtent());
        anchor.setEffectExtent(inline.getEffectExtent());
        anchor.setWrapNone(new CTWrapNone());
        anchor.setDocPr(inline.getDocPr());
        anchor.setCNvGraphicFramePr(inline.getCNvGraphicFramePr());
        anchor.setGraphic(inline.getGraphic());
        return anchor;
    }

    private byte[] readImageBytes() throws Exception {
        final var uri = url.toURI();
        final var scheme = uri.getScheme();
        if ("http".equalsIgnoreCase(scheme) || "https".equalsIgnoreCase(scheme)) {
            try (var client = HttpClient.newBuilder().followRedirects(HttpClient.Redirect.NORMAL).build()) {
                for (int attempt = 0; true; attempt++) {
                    final var request = HttpRequest.newBuilder().uri(uri)
                            .header("User-Agent", "docx4j-builder/1.0").GET().build();
                    final var response = client.send(request, HttpResponse.BodyHandlers.ofByteArray());
                    if (response.statusCode() == 200) {
                        return response.body();
                    }
                    if (response.statusCode() == 429 && attempt < 2) {
                        Thread.sleep(retryAfterMillis(response));
                        continue;
                    }
                    throw new IllegalArgumentException("Failed to download image, status: " + response.statusCode());
                }
            }
        }
        try (final var input = url.openStream()) {
            return input.readAllBytes();
        }
    }

    private static long retryAfterMillis(HttpResponse<?> response) {
        return response.headers().firstValue("Retry-After")
                .map(s -> Math.min(Long.parseLong(s.trim()) * 1000L, 30_000L))
                .orElse(2_000L);
    }

    private boolean isSvg(byte[] bytes) {
        if ("svg".equalsIgnoreCase(FilenameUtils.getExtension(filenameHint))) {
            return true;
        }
        return new String(bytes, 0, Math.min(bytes.length, 512)).contains("<svg");
    }

    private byte[] convertSvgToPng(byte[] svgBytes) throws TranscoderException {
        final var transcoder = new PNGTranscoder();
        if (width != null) {
            transcoder.addTranscodingHint(PNGTranscoder.KEY_WIDTH, width.floatValue());
        }
        if (height != null) {
            transcoder.addTranscodingHint(PNGTranscoder.KEY_HEIGHT, height.floatValue());
        }
        final var input = new TranscoderInput(new ByteArrayInputStream(svgBytes));
        final var out = new ByteArrayOutputStream();
        transcoder.transcode(input, new TranscoderOutput(out));
        return out.toByteArray();
    }

    private void reset() {
        url = null;
        width = null;
        height = null;
        filenameHint = null;
        altText = null;
        horizontalRelativeFrom = null;
        horizontalAlign = null;
        verticalRelativeFrom = null;
        verticalAlign = null;
    }
}
