package com.robin.rapidoffice.elements;

import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;

public class BorderElement {
    protected static final BorderElement NONE = new BorderElement(null, null);

    /**
     * Border style.
     */
    private final String style;

    /**
     * RGB border color.
     */
    private final String rgbColor;


    BorderElement(String style, String rgbColor) {
        this.style = style;
        this.rgbColor = rgbColor;
    }
    void write(String name, XMLWriter w) throws IOException {
        w.append("<").append(name);
        if (style == null && rgbColor == null) {
            w.append("/>");
        } else {
            if (style != null) {
                w.append(" style=\"").append(style).append("\"");
            }
            w.append(">");
            if (rgbColor != null) {
                w.append("<color rgb=\"").append(rgbColor).append("\"/>");
            }
            w.append("</").append(name).append(">");
        }
    }
}
