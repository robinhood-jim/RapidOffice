package com.robin.rapidoffice.meta;

import com.robin.rapidoffice.elements.IWriteableElements;
import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;

public class Fill implements IWriteableElements {
    public static final Fill NONE = new Fill("none", null, true);


    public static final Fill GRAY125 = new Fill("gray125", null, true);
    public static final Fill DARKGRAY = new Fill("darkGray", null, true);
    public static final Fill BLACK = new Fill("black",null , true);


    private final String patternType;

    private final String colorRgb;

    private final boolean fg;
    Fill(String patternType, String colorRgb, boolean fg) {
        this.patternType = patternType;
        this.colorRgb = colorRgb;
        this.fg = fg;
    }
    static Fill fromColor(String fgColorRgb) {
        return fromColor(fgColorRgb, true);
    }
    static Fill fromColor(String colorRgb, boolean fg) {
        return new Fill("solid", colorRgb, fg);
    }

    @Override
    public void writeOut(XMLWriter w) throws IOException {
        w.append("<fill><patternFill patternType=\"").append(patternType).append("\"");
        if (colorRgb == null) {
            w.append("/>");
        } else {
            w.append("><").append(fg ? "fg" : "bg").append("Color rgb=\"").append(colorRgb).append("\"/></patternFill>");
        }
        w.append("</fill>");
    }
}
