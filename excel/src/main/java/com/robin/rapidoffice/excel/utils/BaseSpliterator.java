package com.robin.rapidoffice.excel.utils;

import com.robin.comm.util.xls.ExcelColumnProp;
import com.robin.comm.util.xls.ExcelSheetProp;

import com.robin.core.base.util.Const;
import com.robin.rapidoffice.excel.elements.Cell;
import com.robin.rapidoffice.elements.CellAddress;
import com.robin.rapidoffice.elements.CellType;
import com.robin.rapidoffice.excel.WorkBook;
import com.robin.rapidoffice.exception.ExcelException;
import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.utils.XMLFactoryUtils;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import javax.xml.stream.XMLStreamException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;
import java.util.function.BiFunction;
import java.util.function.Consumer;

abstract class BaseSpliterator<T> implements Spliterator<T> {
    final XMLReader r;
    WorkBook workBook;
    int trackedRowIndex = 0;
    boolean containHeaders=false;
    static ExcelSheetProp prop;
    boolean multipleType=false;
    boolean needIdentifyColumn=false;
    List<Cell> cells=new ArrayList<>();
    boolean finishReadHeader =false;
    Map<Integer,CellAddress> addressMap=new HashMap<>();
    CellProcessor processor;
    static boolean isDate1904;


    public BaseSpliterator(WorkBook workBook, InputStream stream, ExcelSheetProp prop1) throws XMLStreamException {
        this.workBook=workBook;
        isDate1904=workBook.isDate1904();
        this.r =new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),stream);
        prop=prop1;

        processor=new CellProcessor();
        if(prop!=null){
            containHeaders=prop.isFillHeader();
            if(CollectionUtils.isEmpty(prop.getColumnPropList())){
                needIdentifyColumn=true;
            }
        }else{
            containHeaders=true;
            needIdentifyColumn=true;
            prop=ExcelSheetProp.Builder.newBuilder().build();
        }
        r.goTo("sheetData");
    }

    @Override
    public boolean tryAdvance(Consumer<? super T> action) {
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

    @Override
    public Spliterator<T> trySplit() {
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


    boolean hasNext() throws XMLStreamException {
        if(containHeaders && !finishReadHeader){
            readHeader();
            finishReadHeader =true;
        }
        if (r.goTo(() -> r.isStartElement("row") || r.isEndElement("sheetData"))) {
            return "row".equals(r.getLocalName());
        } else {
            return false;
        }
    }
    void readHeader() throws XMLStreamException{
        int trackedColIndex=0;
        while (r.goTo(() -> r.isStartElement("c") || r.isEndElement("row"))) {
            if ("row".equals(r.getLocalName())) {
                break;
            }
            Cell cell = parseCell(trackedColIndex++,false);
            if(needIdentifyColumn && !ObjectUtils.isEmpty(cell.getValue())){
                prop.addColumnProp(new ExcelColumnProp(cell.getValue().toString(),cell.getValue().toString(), Const.META_TYPE_STRING));
            }
        }
        initCells();
    }
    abstract void initCells();
    CellAddress getCellAddressWithFallback(int trackedColIndex) {
        String cellRefOrNull = r.getAttribute("r");
        CellAddress address=addressMap.get(trackedColIndex);
        if(address==null) {
            address = cellRefOrNull != null ? new CellAddress(cellRefOrNull) : new CellAddress(trackedRowIndex, trackedColIndex);
            addressMap.put(trackedColIndex,address);
        }else{
            if(!ObjectUtils.isEmpty(cellRefOrNull)){
                address.setAddress(cellRefOrNull);
            }else{
                address.setPos(trackedRowIndex,trackedColIndex);
            }
        }
        return address;
    }

    T next() throws XMLStreamException {
        if (!"row".equals(r.getLocalName())) {
            throw new NoSuchElementException();
        }
        int trackedColIndex = 0;
        while (r.goTo(() -> r.isStartElement("c") || r.isEndElement("row"))) {
            if ("row".equals(r.getLocalName())) {
                break;
            }
            processCell(trackedColIndex++);
        }
        return constructReturn();
    }

    CellType parseType(String type) {
        switch (type) {
            case "b":
                return CellType.BOOLEAN;
            case "e":
                return CellType.ERROR;
            case "n":
                return CellType.NUMBER;
            case "str":
                return CellType.FORMULA;
            case "s":
            case "inlineStr":
                return CellType.STRING;
        }
        throw new IllegalStateException("Unknown cell type : " + type);
    }

    BiFunction<String, CellAddress,?> getParserForType(CellType type) {
        switch (type) {
            case BOOLEAN:
                return BaseSpliterator::parseBoolean;
            case NUMBER:
                return BaseSpliterator::parseNumber;
            case FORMULA:
            case ERROR:
                return BaseSpliterator::defaultValue;
        }
        throw new IllegalStateException("No parser defined for type " + type);
    }



    private static Object parseNumber(String s,CellAddress address) {
        try {
            if(ObjectUtils.isEmpty(s)){
                return null;
            }
            int column=address.getColumn();
            String tmpVal=s;
            Object retObj=null;
            if(!CollectionUtils.isEmpty(prop.getColumnPropList()) && prop.getColumnPropList().get(column)!=null){
                ExcelColumnProp columnProp=prop.getColumnPropList().get(column);
                switch (columnProp.getColumnType()){
                    case Const.META_TYPE_INTEGER:
                        if(tmpVal.contains(".")){
                            int pos=tmpVal.indexOf(".");
                            tmpVal=tmpVal.substring(0,pos);
                        }
                        retObj=Integer.parseInt(tmpVal);
                        break;
                    case Const.META_TYPE_BIGINT:
                        if(tmpVal.contains(".")){
                            int pos=tmpVal.indexOf(".");
                            tmpVal=tmpVal.substring(0,pos);
                        }
                        retObj=Long.parseLong(tmpVal);
                        break;
                    case Const.META_TYPE_FLOAT:
                        retObj=Float.parseFloat(tmpVal);
                        break;
                    case Const.META_TYPE_DOUBLE:
                    case Const.META_TYPE_NUMERIC:
                        retObj=Double.parseDouble(tmpVal);
                        break;
                    case Const.META_TYPE_TIMESTAMP:
                        retObj=DateUtils.getLocalDateTime(Double.valueOf(s),isDate1904,false);
                        break;
                    default:
                        retObj=new BigDecimal(s);
                }
            }else{
                retObj=new BigDecimal(s);
            }
            return retObj;
        } catch (NumberFormatException e) {
            throw new ExcelException("Cannot parse number : " + s);
        }
    }
    private static String defaultValue(String s,CellAddress address){
        return s;
    }

    private static Boolean parseBoolean(String s,CellAddress address) {
        if ("0".equals(s)) {
            return Boolean.FALSE;
        } else if ("1".equals(s)) {
            return Boolean.TRUE;
        } else {
            throw new ExcelException("Invalid boolean cell value: '" + s + "'. Expecting '0' or '1'.");
        }
    }
    Cell parseCell(int trackedColIndex,boolean isMultiplex) throws XMLStreamException {
        CellAddress addr = getCellAddressWithFallback(trackedColIndex);

        processor.setValue(r,workBook,addr);

        if ("inlineStr".equals(processor.getType())) {
            return parseInlineStr(addr,processor, isMultiplex);
        } else if ("s".equals(processor.getType())) {
            return parseString(addr,processor, isMultiplex);
        } else {
            return parseOther(addr, processor, isMultiplex);
        }
    }

    Cell parseInlineStr(CellAddress addr,CellProcessor processor,boolean isMultiplex) throws XMLStreamException {
        while (r.goTo(() -> r.isStartElement("is") || r.isEndElement("c") || r.isStartElement("f"))) {
            if ("is".equals(r.getLocalName())) {
                processor.setRawValue(r.getValueUntilEndElement("is"));
                processor.setValue(processor.getRawValue());
            } else if ("f".equals(r.getLocalName())) {
                processor.setFormula(r.getValueUntilEndElement("f"));
            } else {
                break;
            }
        }
        processor.setCellType(processor.getFormula() == null ? CellType.STRING : CellType.FORMULA);
        return  returnCell(isMultiplex,workBook,processor,addr);
    }
    Cell empty(CellAddress addr, CellType type) {
        return new Cell(workBook, type, "", addr, null, "");
    }
    Cell parseString(CellAddress addr,CellProcessor processor,boolean isMultiplex) throws XMLStreamException {
        r.goTo(() -> r.isStartElement("v") || r.isEndElement("c"));
        if (r.isEndElement("c")) {
            return empty(addr, CellType.STRING);
        }
        String v = r.getValueUntilEndElement("v");
        if (v.isEmpty()) {
            return empty(addr, CellType.STRING);
        }
        int index = Integer.parseInt(v);
        String sharedStringValue = workBook.getShardingStrings().getValues().get(index).getValue();
        processor.setValue(sharedStringValue);
        processor.setFormula(null);
        processor.setRawValue(sharedStringValue);
        processor.setCellType(CellType.STRING);
        return  returnCell(isMultiplex,workBook,processor,addr);
    }
    Cell parseOther(CellAddress addr, CellProcessor processor,boolean isMultiplex)
            throws XMLStreamException {
        CellType definedType = parseType(processor.getType());
        BiFunction<String,CellAddress, ?> parser = getParserForType(definedType);


        while (r.goTo(() -> r.isStartElement("v") || r.isEndElement("c") || r.isStartElement("f"))) {
            if ("v".equals(r.getLocalName())) {
                processor.setRawValue(r.getValueUntilEndElement("v"));
                try {
                    processor.setValue("".equals(processor.getRawValue()) ? null : parser.apply(processor.getRawValue(),addr));
                } catch (ExcelException e) {
                    definedType = CellType.ERROR;
                }
            } else if ("f".equals(r.getLocalName())) {
                processor.setFormula(r.getValueUntilEndElement("f"));
            } else {
                break;
            }
        }
        if (processor.getFormula() == null && processor.getValue() == null && definedType == CellType.NUMBER) {
            //return new Cell(workBook, CellType.EMPTY, null, addr, null, rawValue);
            processor.setCellType(CellType.EMPTY);
            processor.setValue(null);
            processor.setFormula(null);
            return returnCell(isMultiplex,workBook, processor, addr);
        } else {
            processor.setCellType((processor.getFormula() != null) ? CellType.FORMULA : definedType);
            //return new Cell(workBook, cellType, value, addr, formula, rawValue, dataFormatId, dataFormatString);
            return returnTypeCell(isMultiplex,workBook, processor, addr);
        }
    }
    abstract T constructReturn();

    Cell returnCell(boolean isMultiplex,WorkBook workBook,CellProcessor processor, CellAddress addr) {
        if(!isMultiplex || valueAllEmpty() || cells.get(addr.getColumn())==null) {
            return new Cell(workBook, processor, addr);
        }else{
            Cell cell=cells.get(addr.getColumn());
            if(processor.getValue()!=null){
                cell.setValue(processor.getValue());
            }
            return cell;
        }
    }

    Cell returnTypeCell(boolean isMultiplex,WorkBook workBook, CellProcessor processor, CellAddress addr) {
        if(!isMultiplex || valueAllEmpty() || cells.get(addr.getColumn())==null) {
            return new Cell(workBook, processor.getCellType(), processor.getValue(), addr, processor.getFormula(), processor.getRawValue(),processor.getDataFormatId(),processor.getDataFormatString());
        }else{
            Cell cell=cells.get(addr.getColumn());
            if(processor.getValue()!=null){
                cell.setValue(processor.getValue());
            }
            return cell;
        }
    }
    boolean valueAllEmpty(){
        return CollectionUtils.isEmpty(cells) || cells.stream().allMatch(ObjectUtils::isEmpty);
    }
    abstract void processCell(int trackedColIndex) throws XMLStreamException;

}
