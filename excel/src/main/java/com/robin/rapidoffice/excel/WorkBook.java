package com.robin.rapidoffice.excel;

import cn.hutool.core.io.FileUtil;
import com.robin.comm.util.xls.ExcelSheetProp;
import com.robin.core.fileaccess.util.ByteBufferInputStream;
import com.robin.core.fileaccess.util.ByteBufferOutputStream;
import com.robin.rapidoffice.elements.SheetVisibility;
import com.robin.rapidoffice.elements.StyleHolder;
import com.robin.rapidoffice.excel.elements.Row;
import com.robin.rapidoffice.excel.utils.MapSpliterator;
import com.robin.rapidoffice.excel.utils.RowSpliterator;
import com.robin.rapidoffice.exception.ExcelException;
import com.robin.rapidoffice.meta.Font;
import com.robin.rapidoffice.meta.RelationShip;
import com.robin.rapidoffice.meta.ShardingString;
import com.robin.rapidoffice.meta.ShardingStrings;
import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.utils.OPCPackage;
import com.robin.rapidoffice.utils.ThrowableConsumer;
import com.robin.rapidoffice.utils.XMLFactoryUtils;
import com.robin.rapidoffice.utils.ZipStreamEntry;
import com.robin.rapidoffice.writer.XMLWriter;
import org.apache.commons.io.IOUtils;
import org.apache.flink.core.memory.MemorySegment;
import org.apache.flink.core.memory.MemorySegmentFactory;
import org.springframework.util.Assert;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.Consumer;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


public class WorkBook implements Closeable {
    int activeTab = 0;
    boolean finished = false;
    String applicationName="rapidoffice";
    String applicationVersion;
    OPCPackage opcPackage;
    List<WorkSheet> sheets=new ArrayList<>();
    Map<String, WorkSheet> sheetMap=new HashMap<>();
    boolean date1904;
    ShardingStrings shardingStrings;
    Map<Integer,OutputStream> sheetTmpStreamMap=new HashMap<>();
    Map<Integer,MemorySegment> segmentMap=new HashMap<>();
    Map<Integer,XMLWriter> sheetWriterMap=new HashMap<>();
    XMLWriter writer;
    static String sheetIdPrefix="rId";
    int maxSheetSize;
    String localTmpPath=null;
    final com.robin.rapidoffice.meta.Properties properties = new com.robin.rapidoffice.meta.Properties ();
    boolean readWriteTag=false;
    File file;
    StyleHolder holder=new StyleHolder();
    Map<String, ShardingString> shardingStringMap=new HashMap<>();
    String workBookPath;
    String shardingStringsPath;
    String stylePath;
    String appPath;
    Map<String, String> formatMap =new HashMap<>();
    List<String> formatIdList=new ArrayList<>();
    Map<String,RelationShip> relationShipTypeMap=new HashMap<>();


    public WorkBook(File file,String applicationName,String applicationVersion,int bufferSize){
        this.applicationName=applicationName;
        this.applicationVersion=applicationVersion;
        opcPackage=OPCPackage.create(file,bufferSize);
    }

    public WorkBook(OutputStream cout,String applicationName,String applicationVersion,int bufferSize){
        this.applicationName=applicationName;
        this.applicationVersion=applicationVersion;
        opcPackage=OPCPackage.create(cout,bufferSize);
    }
    public WorkBook(InputStream inputStream) throws XMLStreamException,IOException{
        this(inputStream, ZipStreamEntry.InputStreamBufferMode.HEAP);
    }
    public WorkBook(InputStream inputStream, ZipStreamEntry.InputStreamBufferMode mode) throws XMLStreamException,IOException{
        Assert.notNull(inputStream,"");
        opcPackage=new OPCPackage(inputStream,"UTF8",mode);
        opcPackage.doReadInit((zipFile, zipStreams)->{
            try {
                extractParts();
                extractStyle(stylePath);
                opcPackage.extractRelationShip(OPCPackage.relsNameFor(workBookPath),"_rel");
                extractExtendProperty(appPath);
            }catch (IOException|XMLStreamException ex){
                throw new ExcelException(ex.getMessage());
            }
        });
        beginRead();
    }

    public WorkBook(File file) throws XMLStreamException,IOException {
        if(!FileUtil.exist(file)){
            throw new IOException("file not found!");
        }
        this.file=file;
        opcPackage=OPCPackage.open(file);
        opcPackage.doReadInit((zipFile, zipStreams)->{
            try {
                extractParts();
                extractStyle(stylePath);
                opcPackage.extractRelationShip(OPCPackage.relsNameFor(workBookPath),"_rel");
                extractExtendProperty(appPath);
            }catch (IOException|XMLStreamException ex){
                throw new ExcelException(ex.getMessage());
            }
        });
        beginRead();
    }
    public WorkBook(File path,int bufferSize){
        this.file=path;
        opcPackage=OPCPackage.create(path,bufferSize);
        shardingStrings=new ShardingStrings();
        writer=new XMLWriter(opcPackage.getZipOutStream());
        readWriteTag=true;
    }
    public WorkBook(File path,int bufferSize,int maxSheetSize){
        this.file=path;
        opcPackage=OPCPackage.create(path,bufferSize);
        shardingStrings=new ShardingStrings();
        writer=new XMLWriter(opcPackage.getZipOutStream());
        readWriteTag=true;
        if(maxSheetSize>0) {
            this.maxSheetSize = maxSheetSize;
        }
    }
    public WorkBook(OutputStream outputStream){
        Assert.notNull(outputStream,"");
        opcPackage=OPCPackage.create(outputStream,0);
        writer=new XMLWriter(opcPackage.getZipOutStream());
        shardingStrings=new ShardingStrings();
        readWriteTag=true;
    }
    public WorkBook(OutputStream outputStream,int maxSheetSize){
        Assert.notNull(outputStream,"");
        opcPackage=OPCPackage.create(outputStream,0);
        writer=new XMLWriter(opcPackage.getZipOutStream());
        shardingStrings=new ShardingStrings();
        readWriteTag=true;
        this.maxSheetSize=maxSheetSize;
    }

    public List<String> getFormats(){
        return formatIdList;
    }
    public Map<String,String> getNumFmtMap(){
        return formatMap;
    }

    private void beginRead() throws XMLStreamException,IOException {
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(workBookPath))){
            while(reader.goTo(() -> reader.isStartElement("sheets") || reader.isStartElement("workbookPr") ||
                    reader.isStartElement("workbookView") || reader.isEndElement("workbook"))){
                if ("workbookView".equals(reader.getLocalName())) {
                    String activeTab = reader.getAttribute("activeTab");
                    if (activeTab != null) {
                        this.activeTab = Integer.parseInt(activeTab);
                    }
                } else if ("sheets".equals(reader.getLocalName())) {
                    reader.forEach("sheet", "sheets", this::parseSheet);
                } else if ("workbookPr".equals(reader.getLocalName())) {
                    String date1904Value = reader.getAttribute("date1904");
                    date1904 = Boolean.parseBoolean(date1904Value);
                } else {
                    break;
                }
            }
        }
        shardingStrings=ShardingStrings.formInputStream(opcPackage.getRequiredEntryContent(shardingStringsPath));
    }
    private void beginFlush() throws IOException{
        Assert.notNull(opcPackage.getZipOutStream(),"");
        writeFileContent("[Content_Types].xml", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>");

            w.append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (WorkSheet ws : sheets) {
                int index = getIndex(ws);
                w.append("<Override PartName=\"/xl/worksheets/sheet").append(index).append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");

            }
            w.append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            w.append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");

            w.append("</Types>");
        });
        writeFileContent("_rels/.rels", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            w.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            w.append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>");
            w.append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>");
            w.append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>");

            w.append("</Relationships>");
        });
        writeFileContent("xl/_rels/workbook.xml.rels", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Target=\"sharedStrings.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"/><Relationship Id=\"rId2\" Target=\"styles.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"/>");
            for (WorkSheet ws : sheets) {
                w.append("<Relationship Id=\"rId").append(getIndex(ws) + 2).append("\" Target=\"worksheets/sheet").append(getIndex(ws)).append(".xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"/>");
            }
            w.append("</Relationships>");
        });
        writeFileContent("xl/sharedStrings.xml",shardingStrings::writeOut);
        writeProperty();
        writeWorkBookFile();
        writeFileContent("xl/styles.xml",holder::writeOut);

    }
    void writeProperty() throws IOException{
        writeFileContent("docProps/app.xml", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
            w.append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
            w.append("<Application>").appendEscaped(applicationName).append("</Application>");
            String manager = properties.getManager();
            if (manager != null) {
                w.append("<Manager>");
                w.appendEscaped(manager);
                w.append("</Manager>");
            }
            String company = properties.getCompany();
            if (company != null) {
                w.append("<Company>");
                w.appendEscaped(company);
                w.append("</Company>");
            }
            String hyperlinkBase = properties.getHyperlinkBase();
            if (hyperlinkBase != null) {
                w.append("<HyperlinkBase>");
                w.appendEscaped(hyperlinkBase);
                w.append("</HyperlinkBase>");
            }
            if (applicationVersion != null) {
                w.append("<AppVersion>");
                w.appendEscaped(applicationVersion);
                w.append("</AppVersion>");
            }
            w.append("</Properties>");
        });
        writeFileContent("docProps/core.xml", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>");
            w.append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
            String title = properties.getTitle();
            if (title != null) {
                w.append("<dc:title>");
                w.appendEscaped(title);
                w.append("</dc:title>");
            }
            String subject = properties.getSubject();
            if (subject != null) {
                w.append("<dc:subject>");
                w.appendEscaped(subject);
                w.append("</dc:subject>");
            }
            w.append("<dc:creator>");
            w.appendEscaped(applicationName);
            w.append("</dc:creator>");
            String keywords = properties.getKeywords();
            if (keywords != null) {
                w.append("<cp:keywords>");
                w.appendEscaped(keywords);
                w.append("</cp:keywords>");
            }
            String description = properties.getDescription();
            if (description != null) {
                w.append("<dc:description>");
                w.appendEscaped(description);
                w.append("</dc:description>");
            }
            w.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">");
            w.append(DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss.SSSX").withZone(ZoneId.of("UTC")).format(Instant.now()));
            w.append("</dcterms:created>");
            String category = properties.getCategory();
            if (category != null) {
                w.append("<cp:category>");
                w.appendEscaped(category);
                w.append("</cp:category>");
            }
            w.append("</cp:coreProperties>");
        });
    }
    private void writeWorkbookSheet(XMLWriter w, WorkSheet ws) throws IOException {
        w.append("<sheet name=\"").appendEscaped(ws.getName()).append("\" r:id=\"rId").append(getIndex(ws) + 2)
                .append("\" sheetId=\"").append(getIndex(ws));

        w.append("\"/>");
    }
    void writeWorkBookFile() throws IOException{
        writeFileContent("xl/workbook.xml", w -> {
            w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<workbook " +
                    "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<workbookPr date1904=\"false\"/>" +
                    "<bookViews>" +
                    "<workbookView activeTab=\"" + activeTab + "\"/>" +
                    "</bookViews>" +
                    "<sheets>");

            for (WorkSheet ws : sheets) {
                writeWorkbookSheet(w, ws);
            }
            w.append("</sheets>");

            //w.append("<definedNames>");

            //w.append("</definedNames>");
            w.append("</workbook>");
        });
    }
    ZipEntry beginPart(String partName) throws IOException{
        ZipEntry entry=new ZipEntry(partName);
        if(opcPackage.getZipOutStream()!=null){
            opcPackage.getZipOutStream().putNextEntry(entry);
        }
        return entry;
    }

    void writeFileContent(String partName, ThrowableConsumer<XMLWriter> consumer) throws IOException{
        synchronized (opcPackage.getZipOutStream()) {
            beginPart(partName);
            consumer.accept(writer);
            writer.flush();
            opcPackage.getZipOutStream().closeEntry();
        }
    }

    void beginSheetWrite(WorkSheet sheet,ExcelSheetProp prop) throws IOException{
        int id=sheet.getIndex();
        OutputStream outputStream;
        if(prop.isUseOffHeap()){
            MemorySegment segment= MemorySegmentFactory.allocateOffHeapUnsafeMemory(maxSheetSize);
            segmentMap.put(id,segment);
            outputStream=new ByteBufferOutputStream(segment.getOffHeapBuffer());
        }else{
            if(localTmpPath==null){
                localTmpPath=System.getProperty("java.io.tmpdir")+File.separator+System.currentTimeMillis();
                FileUtil.mkdir(localTmpPath);
            }
            String localSheetPath=localTmpPath+File.separator+"sheet"+id+".xml";
            outputStream=new FileOutputStream(localSheetPath);
        }
        sheetTmpStreamMap.put(id,outputStream);
        sheetWriterMap.put(id,new XMLWriter(outputStream));
        sheet.writeHeader(sheetWriterMap.get(id));
        if(prop.isFillHeader()){
            sheet.writeTitle(sheetWriterMap.get(id),prop);
        }
    }
    int getIndex(WorkSheet ws) {
        synchronized (sheets) {
            return sheets.indexOf(ws) + 1;
        }
    }
    private void parseSheet(XMLReader r){
        String name = r.getAttribute("name");
        String id = r.getAttribute("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        String sheetId = r.getAttribute("sheetId");
        SheetVisibility sheetVisibility;
        if ("veryHidden".equals(r.getAttribute("state"))) {
            sheetVisibility = SheetVisibility.VERY_HIDDEN;
        } else if ("hidden".equals(r.getAttribute("state"))) {
            sheetVisibility = SheetVisibility.HIDDEN;
        } else {
            sheetVisibility = SheetVisibility.VISIBLE;
        }
        int index = sheets.size();
        WorkSheet sheet=new WorkSheet(this,index,id,sheetId,name,sheetVisibility);
        sheets.add(sheet);
        sheetMap.put(name,sheet);
    }
    public WorkSheet createSheet(String sheetName, ExcelSheetProp prop) throws IOException{
        return createSheet(sheetName,prop,null);
    }
    public WorkSheet createSheet(String sheetName, ExcelSheetProp prop, Consumer<WorkSheet> consumer) throws IOException{
        int idx=sheets.size();
        idx++;
        WorkSheet sheet=new WorkSheet(this,prop,idx,sheetIdPrefix+idx,sheetIdPrefix+idx,sheetName,SheetVisibility.VISIBLE);
        sheet.setDefaultStyles(consumer);
        sheets.add(sheet);
        sheetMap.put(sheetName,sheet);
        beginSheetWrite(sheet,prop);
        return sheet;
    }
    public WorkSheet getSheet(String sheetName){
        if(sheetMap.containsKey(sheetName)){
            return sheetMap.get(sheetName);
        }
        return null;
    }

    public Stream<Row> openStream(WorkSheet sheet, ExcelSheetProp prop) throws IOException{
        try{
            InputStream inputStream=getSheetContent(sheet);
            Stream<Row> stream = StreamSupport.stream(new RowSpliterator(this, inputStream,prop), false);
            return stream.onClose(asUncheckedRunnable(inputStream));
        }catch (XMLStreamException ex){
            throw new IOException(ex);
        }
    }
    public InputStream getSheetContent(WorkSheet sheet) throws IOException {
        RelationShip ship = opcPackage.getRelationShipMap().get(sheet.getId());
        if (ship==null || ship.getTarget() == null) {
            String msg = String.format("Sheet#%s '%s' is missing an entry in workbook rels (for id: '%s')",
                    sheet.getIndex(), sheet.getName(), sheet.getId());
            throw new IOException(msg);
        }
        return opcPackage.getRequiredEntryContent(ship.getTarget());
    }
    public Stream<Map<String,Object>> openMapStream(WorkSheet sheet,ExcelSheetProp prop) throws IOException{
        try{
            InputStream inputStream=getSheetContent(sheet);
            Stream<Map<String,Object>> stream = StreamSupport.stream(new MapSpliterator(this, inputStream,prop), false);
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

    public void finish() throws IOException{
        if(finished){
            return;
        }
        if(CollectionUtils.isEmpty(sheets)){
            throw new IllegalArgumentException("A workbook must contain at least one worksheet.");
        }
        if(readWriteTag) {
            beginFlush();
            for (WorkSheet ws : sheets) {
                ws.finish();
                ZipOutputStream outputStream = opcPackage.getZipOutStream();
                beginPart("xl/worksheets/sheet" + ws.getIndex() + ".xml");
                if (ws.prop.isUseOffHeap()) {
                    ByteBufferOutputStream out1 = (ByteBufferOutputStream) sheetTmpStreamMap.get(ws.getIndex());
                    out1.getByteBuffer().position(0);
                    try (ByteBufferInputStream inputStream = new ByteBufferInputStream(out1.getByteBuffer(), out1.getCount())) {
                        IOUtils.copy(inputStream, outputStream, 8192);
                    } finally {
                        out1.close();
                        segmentMap.get(ws.getIndex()).free();
                    }
                } else {
                    try (FileInputStream inputStream = new FileInputStream(localTmpPath + File.separator + "sheet" + ws.getIndex() + ".xml")) {
                        IOUtils.copy(inputStream, outputStream, 8192);
                    } finally {
                        FileUtil.del(localTmpPath + File.separator + "sheet" + ws.getIndex() + ".xml");
                    }
                }
                outputStream.closeEntry();
            }
            if (!sheets.get(0).prop.isUseOffHeap()) {
                FileUtil.del(localTmpPath);
            }
        }
        finished=true;
    }


    void extractParts() throws XMLStreamException,IOException{
        final String contentTypesXml = "[Content_Types].xml";
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(contentTypesXml))){
            while (reader.goTo(() -> reader.isStartElement("Override"))) {
                String contentType = reader.getAttributeRequired("ContentType");
                if(OPCPackage.WORKBOOK_MAIN_CONTENT_TYPE.equals(contentType) || OPCPackage.WORKBOOK_EXCEL_MACRO_ENABLED_MAIN_CONTENT_TYPE.equals(contentType)){
                    workBookPath=reader.getAttributeRequired("PartName");
                }else if(OPCPackage.SHARED_STRINGS_CONTENT_TYPE.equals(contentType)){
                    shardingStringsPath= reader.getAttributeRequired("PartName");
                }else if(OPCPackage.STYLE_CONTENT_TYPE.equals(contentType)){
                    stylePath=reader.getAttributeRequired("PartName");
                }else if(OPCPackage.EXTEND_PROPERTY_CONTENTTYPE.equals(contentType)){
                    appPath=reader.getAttributeRequired("PartName");
                }
                if(workBookPath!=null && shardingStringsPath!=null && stylePath!=null && appPath!=null){
                    break;
                }
            }
            if(workBookPath==null){
                workBookPath="/xl/workbook.xml";
            }
        }
    }
    void extractStyle(String stylePath) throws XMLStreamException,IOException{
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(stylePath))){
            AtomicBoolean insideCellXfs = new AtomicBoolean(false);
            while (reader.goTo(() -> reader.isStartElement("numFmt") || reader.isStartElement("xf") ||
                    reader.isStartElement("cellXfs") || reader.isEndElement("cellXfs"))) {
                if (reader.isStartElement("cellXfs")) {
                    insideCellXfs.set(true);
                } else if (reader.isEndElement("cellXfs")) {
                    insideCellXfs.set(false);
                }
                if ("numFmt".equals(reader.getLocalName())) {
                    String formatCode = reader.getAttributeRequired("formatCode");
                    formatMap.put(reader.getAttributeRequired("numFmtId"), formatCode);
                } else if (insideCellXfs.get() && reader.isStartElement("xf")) {
                    String numFmtId = reader.getAttribute("numFmtId");
                    formatIdList.add(numFmtId);
                    if (OPCPackage.IMPLICIT_NUM_FMTS.containsKey(numFmtId)) {
                        formatMap.put(numFmtId, OPCPackage.IMPLICIT_NUM_FMTS.get(numFmtId));
                    }
                }
            }
        }
    }


    void extractExtendProperty(String extendPropPath) throws XMLStreamException,IOException{
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),opcPackage.getRequiredEntryContent(extendPropPath))){
            while (reader.goTo(()->reader.isStartElement("Application") ||reader.isStartElement("AppVersion"))){
                if("Application".equals(reader.getLocalName())){
                    applicationName=reader.getValueUntilEndElement("Application");
                }else if("AppVersion".equals(reader.getLocalName())){
                    applicationVersion=reader.getValueUntilEndElement("AppVersion");
                }
            }
        }
    }
    public ShardingStrings getShardingStrings(){
        return shardingStrings;
    }

    public boolean isDate1904() {
        return date1904;
    }
    public Optional<WorkSheet> getSheet(int index) {
        return index < 0 || index >= sheets.size() ? Optional.empty() : Optional.of(sheets.get(index));
    }
    public int getSheetNum(){
        return !ObjectUtils.isEmpty(sheets)?sheets.size():0;
    }

    public WorkSheet getFirstSheet() {
        return sheets.get(0);
    }

    public Optional<WorkSheet> findSheet(String name) {
        return sheets.stream().filter(sheet -> name.equals(sheet.getName())).findFirst();
    }

    @Override
    public void close() throws IOException {
        finish();
        if(opcPackage!=null){
            opcPackage.close();
        }
    }
    public void setGlobalDefaultFont(String fontName, double fontSize) {
        this.setGlobalDefaultFont(Font.build(null, null, null, fontName, BigDecimal.valueOf(fontSize), null, null));
    }
    public void setGlobalDefaultFont(Font font) {
        Font.DEFAULT = font;
        this.holder.replaceDefaultFont(font);
    }
    ShardingString addShardingString(String value){
        ShardingString s1=shardingStringMap.get(value);
        if(s1==null){
            s1=new ShardingString(value,shardingStrings.getValues().size());
            shardingStrings.getValues().add(s1);
            shardingStringMap.put(value,s1);
        }
        return s1;
    }


    public static class Builder{
        private static Builder builder;
        private File file;
        private boolean readWriteTag=false;
        private String applicationName;
        private String applicationVersion;

        private InputStream inputStream;
        private OutputStream outputStream;
        private ExcelSheetProp prop;
        private int bufferSize=0;

        private Builder(){

        }
        public static Builder newBuilder(){
            builder=new Builder();
            return builder;
        }
        public Builder readWithInputStream(InputStream stream){
            this.inputStream=stream;
            return this;
        }
        public Builder readWithFile(File file){
            this.file=file;
            return this;
        }
        public Builder applicationName(String applicationName){
            this.applicationName=applicationName;
            return this;
        }
        public Builder applicationVersion(String applicationVersion){
            this.applicationVersion=applicationVersion;
            return this;
        }
        public Builder writeFile(String fileName){
            this.file=new File(fileName);
            this.readWriteTag=true;
            return this;
        }
        public Builder writeFile(File file){
            this.file=file;
            this.readWriteTag=true;
            return this;
        }
        public Builder writeOutputStream(OutputStream outputStream){
            this.outputStream=outputStream;
            this.readWriteTag=true;
            return this;
        }

        public Builder bufferSize(int bufferSize){
            this.bufferSize=bufferSize;
            return this;
        }
        public WorkBook build() throws IOException,XMLStreamException{
            if(!readWriteTag){
                if(file!=null){
                    return new WorkBook(file);
                }else{
                    return new WorkBook(inputStream);
                }
            }else{
                if(file!=null){
                    return new WorkBook(file,applicationName,applicationVersion,bufferSize);
                }else{
                    return new WorkBook(outputStream,applicationName,applicationVersion,bufferSize);
                }
            }
        }
    }

}
