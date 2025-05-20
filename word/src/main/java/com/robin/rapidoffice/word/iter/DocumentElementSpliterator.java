package com.robin.rapidoffice.word.iter;

import com.robin.rapidoffice.exception.ExcelException;
import com.robin.rapidoffice.exception.StreamReadException;
import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.utils.XMLFactoryUtils;
import com.robin.rapidoffice.word.Document;
import com.robin.rapidoffice.word.elements.*;
import org.springframework.util.CollectionUtils;

import javax.xml.stream.XMLStreamException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Spliterator;
import java.util.function.Consumer;

public class DocumentElementSpliterator implements Spliterator<IBodyElement> {
    private final Document document;
    private final InputStream inputStream;
    final XMLReader r;
    String paragraphId;
    String rsidRDefault;
    String rsidR;
    String runId;
    boolean enterRun=false;
    public static final String NAMESPACEMAIN="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public static final String NAMESPACEW14="http://schemas.microsoft.com/office/word/2010/wordml";
    StringBuilder builder=new StringBuilder();
    List<PictureData> pictureDatas=new ArrayList<>();
    IBodyElement element=null;
    List<XRunElement> elements;

    public DocumentElementSpliterator(Document document, InputStream stream) throws XMLStreamException {
        this.document=document;
        this.inputStream=stream;
        r=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),inputStream);
        r.goTo("body");
    }

    @Override
    public boolean tryAdvance(Consumer<? super IBodyElement> action) {
        try {
            if (hasNext()) {
                action.accept(next());
                return true;
            } else {
                return false;
            }
        } catch (XMLStreamException e) {
            throw new ExcelException(e);
        }
    }
    boolean hasNext() throws XMLStreamException {

        if(pictureDatas.size()>0){
            pictureDatas.clear();
        }
        element=null;
        elements=new ArrayList<>();
        while(r.hasNext()) {
            if (r.isStartElement("p")) {
                paragraphId = r.getAttribute(NAMESPACEW14, "paraId");
                rsidRDefault=r.getAttribute(NAMESPACEMAIN,"rsidRDefault");
                rsidR=r.getAttribute(NAMESPACEMAIN,"rsidR");
                r.goTo(() -> r.isStartElement("r") || r.isStartElement("t"));
            }
            if(r.isEndElement("p")){
                element = new ParagraphElement(paragraphId,rsidR,rsidRDefault, runId, elements);
                r.next();
                break;
            }
            if (r.isStartElement("r")) {
                enterRun = true;
                runId = null;
                if (r.getAttributeCount() > 0) {
                    runId = r.getAttributeAt(0);
                }
                XRunElement element1 = new XRunElement();
                if(builder.length()>0){
                    builder.delete(0,builder.length());
                }
                r.doUntilEndElement("r", r2 -> {
                    String name = r.getLocalName();
                    switch (name) {
                        case "t":
                            builder.append(r.getValueUntilEndElement("t"));
                            break;
                        case "blip":
                            PictureData pictureData1 = new PictureData();
                            pictureData1.setRId(r2.getAttributeValue(0));
                            pictureData1.setPath(document.getOpcPackage().getRelationShipMap().get(r2.getAttributeValue(0)).getTarget());
                            pictureDatas.add(pictureData1);
                            break;
                    }
                });
                element1.setId(runId);
                if (builder.length() > 0) {
                    element1.setContent(builder.toString());
                }
                if (pictureDatas.size() > 0) {
                    element1.setPictureDatas(pictureDatas);
                }
                elements.add(element1);
            }else if(r.isStartElement("tbl")){
                TableElement table=parseTbl();
                if(table!=null){
                    element=table;
                    break;
                }
            } else{
                r.next();
            }
        }
        return element!=null;
    }
    private TableElement parseTbl(){
        List<String> header=new ArrayList<>();
        List<String> tmpValue=new ArrayList<>();
        List<List<String>> values=new ArrayList<>();
        StringBuilder builder1=new StringBuilder();
        boolean firstRow=true;
        try {
            if (r.isStartElement("tbl")) {
                while (r.hasNext()) {
                    r.next();
                    if (r.isStartElement("tr")) {
                        if (tmpValue.size() > 0) {
                            tmpValue=new ArrayList<>();
                        }
                    }
                    if (r.isStartElement("tc")) {
                        if (builder1.length() > 0) {
                            builder1.delete(0, builder1.length());
                        }
                    } else if (r.isStartElement("t")) {
                        builder1.append(r.getValueUntilEndElement("t"));
                    } else if (r.isEndElement("tc")) {
                        tmpValue.add(builder1.toString());
                    } else if (r.isEndElement("tr")) {
                        if (firstRow) {
                            header.addAll(tmpValue);
                            firstRow = false;
                        } else {
                            values.add(tmpValue);
                        }
                    } else if (r.isEndElement("tbl")) {
                        r.next();
                        break;
                    }
                }
            }
            if(!CollectionUtils.isEmpty(values)){
                return new TableElement(header,values);
            }else{
                return null;
            }
        }catch (XMLStreamException ex){
            throw new StreamReadException(ex);
        }
    }
    public IBodyElement next() throws XMLStreamException{
        return element;
    }

    @Override
    public Spliterator<IBodyElement> trySplit() {
        return null;
    }

    @Override
    public long estimateSize() {
        return Long.MAX_VALUE;
    }

    @Override
    public int characteristics() {
        return DISTINCT | IMMUTABLE | NONNULL | ORDERED;
    }
}
