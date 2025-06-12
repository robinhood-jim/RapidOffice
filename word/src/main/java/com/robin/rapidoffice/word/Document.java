package com.robin.rapidoffice.word;

import cn.hutool.core.io.FileUtil;
import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.utils.OPCPackage;
import com.robin.rapidoffice.utils.XMLFactoryUtils;
import com.robin.rapidoffice.word.elements.*;
import com.robin.rapidoffice.word.iter.DocumentElementSpliterator;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class Document implements Closeable {
    List<Footer> footers=new ArrayList<>();
    List<Header> headers=new ArrayList<>();
    List<Theme> themes=new ArrayList<>();
    private File file;
    InputStream inputStream;
    private OPCPackage opcPackage;
    private String documentPath;
    private String fontTablePath;
    private String stylePath;
    private List<String> headerPaths=new ArrayList<>();
    List<String> footerPaths=new ArrayList<>();
    List<String> themePaths=new ArrayList<>();
    PrDefaultRpr defaultRpr=new PrDefaultRpr();

    List<LsdException> lsdExceptions=new ArrayList<>();


    public Document(File file) throws IOException, XMLStreamException {
        if(!FileUtil.exist(file)){
            throw new IOException("file not found!");
        }
        this.file=file;
        opcPackage= OPCPackage.open(file);
        opcPackage.doReadInit((zipFile,zipStrems)->{
            try{
                extractParts();
                extractStyle();
                opcPackage.extractRelationShip(OPCPackage.relsNameFor(documentPath),"_rel");
            }catch (IOException|XMLStreamException ex){

            }
        });
    }
    public Document(InputStream inputStream) throws IOException{
        this.inputStream=inputStream;
        opcPackage=OPCPackage.open(inputStream);
        opcPackage.doReadInit((zipFile,zipStrems)->{
            try{
                extractParts();
                extractStyle();
                opcPackage.extractRelationShip(OPCPackage.relsNameFor(documentPath),"_rel");
            }catch (IOException|XMLStreamException ex){

            }
        });
    }

    void extractParts() throws IOException,XMLStreamException{
        final String contentTypesXml = "[Content_Types].xml";
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(contentTypesXml))){
            while (reader.goTo(() -> reader.isStartElement("Override"))) {
                String contentType = reader.getAttributeRequired("ContentType");
                if(OPCPackage.WORD_DOCUMENT_CONTENT_TYPE.equals(contentType)){
                    documentPath=reader.getAttributeRequired("PartName");
                }else if(OPCPackage.WORD_FONTTABLE_CONTENT_TYPE.equals(contentType)){
                    fontTablePath= reader.getAttributeRequired("PartName");
                }else if(OPCPackage.WORD_STYLE_CONTENT_TYPE.equals(contentType)){
                    stylePath=reader.getAttributeRequired("PartName");
                }else if(OPCPackage.WORD_HEADER_CONTENT_TYPE.equals(contentType)){
                    headerPaths.add(reader.getAttributeRequired("PartName"));
                }else if(OPCPackage.WORD_FOOTER_CONTENT_TYPE.equals(contentType)){
                    footerPaths.add(reader.getAttributeRequired("PartName"));
                }else if(OPCPackage.WORD_THEME_CONTENT_TYPE.equals(contentType)){
                    themePaths.add(reader.getAttributeRequired("PartName"));
                }
            }
            if(documentPath==null){
                documentPath="/xl/document.xml";
            }
        }
    }
    void extractStyle() throws XMLStreamException,IOException{
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(stylePath))){
            reader.goTo("docDefaults");
            while (reader.goTo("rPr")){
                while(reader.goTo("rFonts")){
                    reader.doInAttributes(r->{
                        for(int i=0;i<r.getAttributeCount();i++){
                            defaultRpr.getFonts().put(r.getAttributeName(i).getLocalPart(),r.getAttributeValue(i));
                        }
                    });
                }
                while(reader.goTo("lang")){
                    reader.doInAttributes(r->{
                        for(int i=0;i<r.getAttributeCount();i++) {
                            defaultRpr.getLangs().put(r.getAttributeName(i).getLocalPart(),r.getAttributeValue(i));
                        }
                    });
                }
            }
            reader.goTo("latentStyles");
            while(reader.goTo("lsdException")){
                reader.doInAttributes(r->{
                    LsdException.Builder builder=LsdException.Builder.newBuilder();
                    for(int i=0;i<r.getAttributeCount();i++) {
                        String name=r.getAttributeName(i).getLocalPart();
                        switch (name){
                            case "name":
                                builder.name(r.getAttributeValue(i));
                                break;
                            case "qFormat":
                                builder.qFormat(r.getAttributeValue(i));
                                break;
                            case "semiHidden":
                                builder.semiHidden(r.getAttributeValue(i));
                                break;
                            case "unhideWhenUsed":
                                builder.unhideWhenUsed(r.getAttributeValue(i));
                                break;
                            case "uiPriority":
                                builder.uiPriority(r.getAttributeValue(i));
                                break;
                        }
                        lsdExceptions.add(builder.build());
                    }
                });
            }
            while (reader.goTo("style")){
                final Style style=new Style();
                reader.doInAttributes(r->{
                    for(int i=0;i<r.getAttributeCount();i++){
                        String name=r.getAttributeName(i).getLocalPart();
                        switch (name){
                            case "type":
                                style.setType(r.getAttributeValue(i));
                                break;
                            case "default":
                                style.setDefaultVal(r.getAttributeValue(i));
                                break;
                            case "styleId":
                                style.setStyleId(r.getAttributeValue(i));
                                break;
                        }
                    }
                    reader.doUntilEndElement("style", r1 -> {
                        String name=r.getLocalName();
                        switch (name){
                            case "name":
                                style.setName(r.getAttributeValue(0));
                                break;
                            case "baseOn":
                                style.setBaseOn(r.getAttributeValue(0));
                                break;
                            case "qFormat":
                                if(r.getAttributeCount()>0){
                                    style.setQFormat(r.getAttributeValue(0));
                                }
                                break;
                            case "uiPriority":
                                if(r.getAttributeCount()>0){
                                    style.setUiPriority(r.getAttributeValue(0));
                                }
                                break;
                            case "pPr":
                                reader.doUntilEndElement("pPr",r2->{
                                    String name2=r2.getLocalName();
                                    switch (name2){
                                        case "rFonts":
                                            style.getRpr().getFonts().put(r2.getAttributeName(0).getLocalPart(),r.getAttributeValue(0));
                                            break;
                                        case "lang":
                                            style.getRpr().getLangs().put(r2.getAttributeName(0).getLocalPart(),r.getAttributeValue(0));
                                            break;
                                        case "sz":
                                            style.getRpr().setSz(r2.getAttributeValue(0));
                                            break;
                                        case "szCs":
                                            style.getRpr().setSzCs(r2.getAttributeValue(0));
                                            break;
                                    }
                                });
                                break;
                        }
                    });

                });
            }
        }
    }
    public Stream<IBodyElement> openStream() throws IOException{
        try{
            InputStream inputStream=opcPackage.getRequiredEntryContent(documentPath);
            Stream<IBodyElement> stream = StreamSupport.stream(new DocumentElementSpliterator(this, inputStream), false);
            return stream.onClose(asUncheckedRunnable(inputStream));
        }catch (XMLStreamException ex){
            throw new IOException(ex);
        }
    }
    private static Runnable asUncheckedRunnable(Closeable c) {
        return () -> {
            try {
                c.close();
            } catch (IOException e) {
                throw new UncheckedIOException(e);
            }
        };
    }

    public OPCPackage getOpcPackage() {
        return opcPackage;
    }

    @Override
    public void close() throws IOException {
        if(opcPackage!=null){
            opcPackage.close();
        }
        if(inputStream!=null){
            inputStream.close();
        }
    }
}
