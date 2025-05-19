package com.robin.rapidoffice.meta;

import com.robin.rapidoffice.elements.IWriteableElements;
import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Locale;

public class Font implements IWriteableElements {
    static String defaultFontName= Locale.CHINA.equals(Locale.getDefault()) || Locale.SIMPLIFIED_CHINESE.equals(Locale.getDefault()) ? "宋体" : "Calibri";
    public static Font DEFAULT = build(false, false, false, defaultFontName, BigDecimal.valueOf(12.0), "FF000000", false);

    /**
     * Bold flag.
     */
    private final boolean bold;
    /**
     * Italic flag.
     */
    private final boolean italic;
    /**
     * Underlined flag.
     */
    private final boolean underlined;
    /**
     * Font name.
     */
    private final String name;
    /**
     * Font size.
     */
    private final BigDecimal size;
    /**
     * RGB font color.
     */
    private final String rgbColor;
    /**
     * Strikethrough flag.
     */
    private final boolean strikethrough;
    public Font(boolean bold, boolean italic, boolean underlined, String name, BigDecimal size, String rgbColor, boolean strikethrough) {
        if (size.compareTo(BigDecimal.valueOf(409)) > 0 || size.compareTo(BigDecimal.valueOf(1)) < 0) {
            throw new IllegalStateException("Font size must be between 1 and 409 points: " + size);
        }
        this.bold = bold;
        this.italic = italic;
        this.underlined = underlined;
        this.name = name;
        this.size = size.setScale(2, RoundingMode.HALF_UP);
        this.rgbColor = rgbColor;
        this.strikethrough = strikethrough;
    }
    public static Font build(Boolean bold, Boolean italic, Boolean underlined, String name, BigDecimal size, String rgbColor, Boolean strikethrough) {
        return new Font(bold != null? bold : DEFAULT.bold, italic != null ? italic : DEFAULT.italic , underlined != null ? underlined : DEFAULT.underlined, name != null ? name : DEFAULT.name, size != null ?  size:DEFAULT.size, rgbColor != null ?  rgbColor: DEFAULT.rgbColor, strikethrough != null ? strikethrough : DEFAULT.strikethrough);
    }

    @Override
    public void writeOut(XMLWriter w) throws IOException {
        w.append("<font>").append(bold ? "<b/>" : "").append(italic ? "<i/>" : "").append(underlined ? "<u/>" : "").append("<sz val=\"").append(size.toString()).append("\"/>");
        w.append(strikethrough ? "<strike/>" : "");
        if (rgbColor != null) {
            w.append("<color rgb=\"").append(rgbColor).append("\"/>");
        }
        w.append("<name val=\"").append(name).append("\"/>");
        w.append("</font>");
    }
}
