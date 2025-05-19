package com.robin.rapidoffice.elements;

import com.robin.rapidoffice.meta.Fill;
import com.robin.rapidoffice.meta.Font;
import com.robin.rapidoffice.utils.OPCPackage;
import com.robin.rapidoffice.utils.ThrowableConsumer;
import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.function.Function;

public class StyleHolder implements IWriteableElements {
    private final ConcurrentMap<String, Integer> valueFormattings = new ConcurrentHashMap<>();
    private final ConcurrentMap<Font, Integer> fonts = new ConcurrentHashMap<>();
    private final ConcurrentMap<Fill, Integer> fills = new ConcurrentHashMap<>();
    private final ConcurrentMap<Border, Integer> borders = new ConcurrentHashMap<>();
    private final ConcurrentMap<CellStyle, Integer> styles = new ConcurrentHashMap<>();
    public StyleHolder(){
        mergeCellStyle(0, "0", Font.DEFAULT, Fill.NONE, Border.NONE, null);
        cacheFill(Fill.DARKGRAY);
    }

    public void replaceDefaultFont(Font font) {
        fonts.entrySet().removeIf(entry->entry.getValue()==0);
        fonts.putIfAbsent(font,0);
    }
    private static <T> int cacheStuff(ConcurrentMap<T, Integer> cache, T t, Function<T, Integer> indexFunction) {
        return cache.computeIfAbsent(t, indexFunction);
    }
    private static <T> int cacheStuff(ConcurrentMap<T, Integer> cache, T t) {
        return cacheStuff(cache, t, k -> cache.size());
    }
    public int mergeCellStyle(int currentStyle,String numberFormat,Font font,Fill fill,Border border, Alignment alignment){
        CellStyle origin=styles.entrySet().stream().filter(e->e.getValue().equals(currentStyle)).map(Map.Entry::getKey).findFirst().orElse(null);
        CellStyle style=new CellStyle(origin,cachedValueFormat(numberFormat),cacheFont(font),cacheBorder(border),cacheFill(fill),alignment);
        return cacheStuff(styles,style);
    }
    public int mergeCellStyle(int currentStyle,int numberFormatId,Font font,Fill fill,Border border, Alignment alignment){
        CellStyle origin=styles.entrySet().stream().filter(e->e.getValue().equals(currentStyle)).map(Map.Entry::getKey).findFirst().orElse(null);
        CellStyle style=new CellStyle(origin,numberFormatId,cacheFont(font),cacheBorder(border),cacheFill(fill),alignment);
        return cacheStuff(styles,style);
    }
    int cachedValueFormat(String valueFormat){
        Integer numFmtId=OPCPackage.IMPLICIT_NUM_FMTS.entrySet().stream().filter(f->f.getValue().equals(valueFormat)).map(Map.Entry::getKey).map(Integer::valueOf).findFirst().orElse(null);
        if(numFmtId==null){
            numFmtId=cacheStuff(valueFormattings,valueFormat,k->valueFormattings.size()+165);
        }
        return numFmtId;
    }
    int cacheFont(Font f) {
        return cacheStuff(fonts, f);
    }
    int cacheFill(Fill f) {
        return cacheStuff(fills, f);
    }
    int cacheBorder(Border b) {
        return cacheStuff(borders, b);
    }
    private static <T> void writeContent(XMLWriter w, Map<T, Integer> cache, String name, ThrowableConsumer<Map.Entry<T, Integer>> consumer) throws IOException {
        w.append("<").append(name).append(" count=\"").append(cache.size()).append("\">");
        List<Map.Entry<T, Integer>> entries = new ArrayList<>(cache.entrySet());
        entries.sort(Comparator.comparingInt(Map.Entry::getValue));
        for (Map.Entry<T, Integer> e : entries) {
            consumer.accept(e);
        }
        w.append("</").append(name).append(">");
    }

    @Override
    public void writeOut(XMLWriter w) throws IOException {
        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        writeContent(w, valueFormattings, "numFmts", e -> w.append("<numFmt numFmtId=\"").append(e.getValue()).append("\" formatCode=\"").append(e.getKey()).append("\"/>"));
        writeContent(w, fonts, "fonts", e -> e.getKey().writeOut(w));
        //writeContent(w, fills, "fills", e -> e.getKey().writeOut(w));
        writeContent(w, borders, "borders", e -> e.getKey().writeOut(w));
        w.append("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
        writeContent(w, styles, "cellXfs", e -> e.getKey().writeOut(w));

        w.append("</styleSheet>");
    }
}
