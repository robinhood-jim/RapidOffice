package com.robin.rapidoffice.utils;


import com.robin.comm.util.xls.ExcelColumnProp;
import com.robin.core.base.util.Const;
import com.robin.core.base.util.FileUtils;
import com.robin.rapidoffice.elements.CellType;
import com.robin.rapidoffice.meta.RelationShip;

import com.robin.rapidoffice.exception.ExcelException;
import com.robin.rapidoffice.reader.XMLReader;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.springframework.util.Assert;
import org.springframework.util.ObjectUtils;

import javax.xml.stream.XMLStreamException;
import java.io.*;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.regex.Pattern;
import java.util.zip.ZipOutputStream;


@Slf4j
public class OPCPackage implements Closeable {
    private static final Pattern filenameRegex = Pattern.compile("^(.*/)([^/]+)$");
    public static final Map<String, String> IMPLICIT_NUM_FMTS = new HashMap<>() {{
        put("1", "0");
        put("2", "0.00");
        put("3", "#,##0");
        put("4", "#,##0.00");
        put("9", "0%");
        put("10", "0.00%");
        put("11", "0.00E+00");
        put("12", "# ?/?");
        put("13", "# ??/??");
        put("14", "mm-dd-yy");
        put("15", "d-mmm-yy");
        put("16", "d-mmm");
        put("17", "mmm-yy");
        put("18", "h:mm AM/PM");
        put("19", "h:mm:ss AM/PM");
        put("20", "h:mm");
        put("21", "h:mm:ss");
        put("22", "m/d/yy h:mm");
        put("37", "#,##0 ;(#,##0)");
        put("38", "#,##0 ;[Red](#,##0)");
        put("39", "#,##0.00;(#,##0.00)");
        put("40", "#,##0.00;[Red](#,##0.00)");
        put("45", "mm:ss");
        put("46", "[h]:mm:ss");
        put("47", "mmss.0");
        put("48", "##0.0E+0");
        put("49", "@");
    }};
    private ZipFile zipFile;
    private ZipArchiveInputStream zipStreams;

    private ZipOutputStream zipOutStream;
    private BufferedOutputStream bufferedStream;
    private FileOutputStream fileOutputStream=null;
    private static int DEFAULTBUFFEREDSIZE=4096;
    private ZipStreamEntry entry;
    private String workBookPath;
    private String shardingStringsPath;
    private String stylePath;
    private String appPath;
    public static final String WORKBOOK_MAIN_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    public static final String WORKBOOK_EXCEL_MACRO_ENABLED_MAIN_CONTENT_TYPE =
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
    public static final String SHARED_STRINGS_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    public static final String STYLE_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";

    public static final String CORE_PROPERTIY_CONTENTTYPE="application/vnd.openxmlformats-package.core-properties+xml";
    public static final String EXTEND_PROPERTY_CONTENTTYPE="application/vnd.openxmlformats-officedocument.extended-properties+xml";

    private Map<String, String> formatMap =new HashMap<>();
    private List<String> formatIdList=new ArrayList<>();
    private Map<String, RelationShip> relationShipMap=new HashMap<>();
    private Map<String,RelationShip> relationShipTypeMap=new HashMap<>();
    private String applicationName;
    private String appVersion;
    private boolean readWriteMode=false;

    private OPCPackage(File zipFile){
        try{
            this.zipFile =new ZipFile(zipFile);
            readInit();
        }catch (IOException ex){

        }

    }
    private OPCPackage(File targetFile,int bufferedSize) throws IOException{
        readWriteMode=true;
        this.fileOutputStream=new FileOutputStream(targetFile);
        this.bufferedStream=new BufferedOutputStream(fileOutputStream,bufferedSize>0?bufferedSize:DEFAULTBUFFEREDSIZE);
        this.zipOutStream=new ZipOutputStream(fileOutputStream);
    }

    private OPCPackage(OutputStream outputStream,int bufferedSize){
        this.bufferedStream=new BufferedOutputStream(outputStream,bufferedSize>0?bufferedSize:DEFAULTBUFFEREDSIZE);
        readWriteMode=true;
        zipOutStream=new ZipOutputStream(bufferedStream);
    }
    private OPCPackage(InputStream stream,String encode){
        zipStreams=new ZipArchiveInputStream(stream,encode);
        readInit();
    }
    private OPCPackage(InputStream stream){
        this(stream,"UTF8");
    }
    private void readInit() throws ExcelException {
        try {
            if (ObjectUtils.isEmpty(zipFile) && !ObjectUtils.isEmpty(zipStreams)) {
                entry = new ZipStreamEntry(zipStreams);
            }
            extractParts();
            extractStyle(stylePath);
            extractRelationShip(relsNameFor(workBookPath));
            extractExtendProperty(appPath);
        }catch (IOException|XMLStreamException ex){
            throw new ExcelException(ex.getMessage());
        }
    }
    public static OPCPackage open(File zipFile){
        return new OPCPackage(zipFile);
    }
    public static OPCPackage open(InputStream stream){
        return new OPCPackage(stream);
    }
    public static OPCPackage create(File fileName){
        return create(fileName,DEFAULTBUFFEREDSIZE);
    }
    public static OPCPackage create(OutputStream outputStream){
        return create(outputStream,DEFAULTBUFFEREDSIZE);
    }
    public static OPCPackage create(OutputStream outputStream,int bufferedSize){
        try{
            OPCPackage opcPackage=new OPCPackage(outputStream,bufferedSize);
            return opcPackage;
        }catch (Exception ex){
            log.error("{}",ex.getMessage());
        }
        return null;
    }
    public static OPCPackage create(File fileName,int bufferedSize){
        try{
            FileUtils.mkDirReclusive(fileName.getAbsolutePath());
            OPCPackage opcPackage=new OPCPackage(fileName,bufferedSize);
            return opcPackage;

        }catch (IOException ex){

        }
        return null;
    }
    private void extractParts() throws XMLStreamException,IOException{
        final String contentTypesXml = "[Content_Types].xml";
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),getRequiredEntryContent(contentTypesXml))){
            while (reader.goTo(() -> reader.isStartElement("Override"))) {
                String contentType = reader.getAttributeRequired("ContentType");
                if(WORKBOOK_MAIN_CONTENT_TYPE.equals(contentType) ||WORKBOOK_EXCEL_MACRO_ENABLED_MAIN_CONTENT_TYPE.equals(contentType)){
                    workBookPath=reader.getAttributeRequired("PartName");
                }else if(SHARED_STRINGS_CONTENT_TYPE.equals(contentType)){
                    shardingStringsPath= reader.getAttributeRequired("PartName");
                }else if(STYLE_CONTENT_TYPE.equals(contentType)){
                    stylePath=reader.getAttributeRequired("PartName");
                }else if(EXTEND_PROPERTY_CONTENTTYPE.equals(contentType)){
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
    private void extractStyle(String stylePath) throws XMLStreamException,IOException{
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),getRequiredEntryContent(stylePath))){
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
                    if (IMPLICIT_NUM_FMTS.containsKey(numFmtId)) {
                        formatMap.put(numFmtId, IMPLICIT_NUM_FMTS.get(numFmtId));
                    }
                }
            }
        }
    }
    private static String relsNameFor(String entryName) {
        return filenameRegex.matcher(entryName).replaceFirst("$1_rels/$2.rels");
    }
    private void extractRelationShip(String relationPath) throws XMLStreamException,IOException{
        String xlFolder = relationPath.substring(0, relationPath.indexOf("_rel"));
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),getRequiredEntryContent(relationPath))){
            while(reader.goTo("Relationship")){
                String id = reader.getAttribute("Id");
                String target = reader.getAttribute("Target");
                String type=reader.getAttribute("Type");
                // if name does not start with /, it is a relative path
                if (!target.startsWith("/")) {
                    target = xlFolder + target;
                } // else it is an absolute path
                relationShipMap.put(id, new RelationShip(id,target,type));
            }
        }
    }
    private void extractExtendProperty(String extendPropPath) throws XMLStreamException,IOException{
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),getRequiredEntryContent(extendPropPath))){
            while (reader.goTo(()->reader.isStartElement("Application") ||reader.isStartElement("AppVersion"))){
                if("Application".equals(reader.getLocalName())){
                    applicationName=reader.getValueUntilEndElement("Application");
                }else if("AppVersion".equals(reader.getLocalName())){
                    appVersion=reader.getValueUntilEndElement("AppVersion");
                }
            }
        }
    }
    public InputStream getRequiredEntryContent(String name) throws IOException {
        return Optional.ofNullable(getEntryContent(name))
                .orElseThrow(() -> new IOException(name + " not found"));
    }
    private InputStream getEntryContent(String name) throws IOException{
        Assert.notNull(name,"");
        String tname=name;
        if (tname.startsWith("/")) {
            tname = tname.substring(1);
        }
        if(!ObjectUtils.isEmpty(zipFile)){
            ZipArchiveEntry entry1=zipFile.getEntry(tname);
            if(entry1==null){
                Enumeration<ZipArchiveEntry> entries = zipFile.getEntries();
                while (entries.hasMoreElements()) {
                    ZipArchiveEntry e = entries.nextElement();
                    if (e.getName().equalsIgnoreCase(tname)) {
                        return zipFile.getInputStream(e);
                    }
                }
                return null;
            }
            return zipFile.getInputStream(entry1);
        }else if(!ObjectUtils.isEmpty(entry)){
            return entry.getInputStream(tname);
        }else{
            throw new IOException("init incorrect!");
        }

    }
    public InputStream getWorkBookContent() throws IOException{
        return getRequiredEntryContent(workBookPath);
    }

    public InputStream getShardingStrings() throws IOException{
        return getRequiredEntryContent(shardingStringsPath);
    }
    public void finish(){
        if(readWriteMode){

        }
    }


    @Override
    public void close() throws IOException {
        if(zipFile!=null){
            zipFile.close();
        }
        if(entry!=null){
            entry.close();
        }
        if(bufferedStream!=null){
            bufferedStream.flush();
            if(zipOutStream!=null){
                zipOutStream.closeEntry();
                zipOutStream.close();
            }
            if(fileOutputStream!=null){
                fileOutputStream.close();
            }

        }
    }

    public Map<String, RelationShip> getRelationShipMap() {
        return relationShipMap;
    }

    public String getApplicationName() {
        return applicationName;
    }

    public String getAppVersion() {
        return appVersion;
    }
    public List<String> getFormatIdList(){
        return formatIdList;
    }
    public Map<String,String> getFormatMap(){
        return formatMap;
    }

    public ZipFile getZipFile() {
        return zipFile;
    }

    public ZipOutputStream getZipOutStream() {
        return zipOutStream;
    }

    public BufferedOutputStream getBufferedStream() {
        return bufferedStream;
    }
    public static CellType parseCellType(ExcelColumnProp prop){
        CellType type=CellType.EMPTY;
        switch (prop.getColumnType()){
            case Const.META_TYPE_BIGINT:
            case Const.META_TYPE_FLOAT:
            case Const.META_TYPE_DOUBLE:
            case Const.META_TYPE_DECIMAL:
            case Const.META_TYPE_INTEGER:
            case Const.META_TYPE_TIMESTAMP:
                type=CellType.NUMBER;
                break;
            case Const.META_TYPE_BOOLEAN:
                type=CellType.BOOLEAN;
                break;
            case Const.META_TYPE_FORMULA:
                type=CellType.FORMULA;
                break;
            default:
                type=CellType.STRING;
        }
        return type;
    }
}
