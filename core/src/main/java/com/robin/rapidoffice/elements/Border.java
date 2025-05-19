package com.robin.rapidoffice.elements;

import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;
import java.util.EnumMap;
import java.util.Map;

public class Border implements IWriteableElements{
    public static final Border NONE = new Border();
    public static final Border BLACK=new Border(new BorderElement(null,"FFFFFFFF"));


    final Map<BorderSide, BorderElement> elements = new EnumMap<>(BorderSide.class);
    Border() {
        this(BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE, BorderElement.NONE);
    }


    Border(BorderElement element) {
        this(element, element, element, element, BorderElement.NONE);
    }

    
    Border(BorderElement left, BorderElement right, BorderElement top, BorderElement bottom, BorderElement diagonal) {
        elements.put(BorderSide.TOP, top);
        elements.put(BorderSide.LEFT, left);
        elements.put(BorderSide.BOTTOM, bottom);
        elements.put(BorderSide.RIGHT, right);
        elements.put(BorderSide.DIAGONAL, diagonal);
    }

    @Override
    public void writeOut(XMLWriter w) throws IOException {
        w.append("<border");

        w.append(">");
        elements.get(BorderSide.LEFT).write("left", w);
        elements.get(BorderSide.RIGHT).write("right", w);
        elements.get(BorderSide.TOP).write("top", w);
        elements.get(BorderSide.BOTTOM).write("bottom", w);
        elements.get(BorderSide.DIAGONAL).write("diagonal", w);
        w.append("</border>");
    }
}
