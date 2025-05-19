package com.robin.rapidoffice.excel;

import com.robin.comm.util.xls.ExcelSheetProp;
import lombok.extern.slf4j.Slf4j;

import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

@Slf4j
public class SingleWorkBook extends WorkBook {
    private ExcelSheetProp prop;
    private WorkSheet currentSheet;
    private int sheetMaxRows=WorkSheet.MAX_ROWS/4;
    int totalRow=0;

    public SingleWorkBook(File file, String applicationName, String applicationVersion, int bufferSize) {
        super(file, applicationName, applicationVersion, bufferSize);
    }

    public SingleWorkBook(OutputStream cout, String applicationName, String applicationVersion, int bufferSize) {
        super(cout, applicationName, applicationVersion, bufferSize);
    }

    public SingleWorkBook(InputStream inputStream) throws XMLStreamException, IOException {
        super(inputStream);
    }

    public SingleWorkBook(File file) throws XMLStreamException, IOException {
        super(file);
    }

    public SingleWorkBook(File path, int bufferSize, ExcelSheetProp prop) {
        super(path, bufferSize);
        this.prop=prop;
    }
    public SingleWorkBook(File path, int bufferSize, ExcelSheetProp prop,int maxRows) {
        super(path, bufferSize);
        this.prop=prop;
        if(maxRows>0) {
            this.sheetMaxRows = maxRows;
        }
    }

    public SingleWorkBook(OutputStream outputStream, ExcelSheetProp prop) {
        super(outputStream);
        this.prop=prop;
    }
    public SingleWorkBook(OutputStream outputStream, ExcelSheetProp prop,int sheetMaxRows) {
        super(outputStream);
        this.prop=prop;
        this.sheetMaxRows=sheetMaxRows;
    }
    public void beginWrite() throws IOException{
        int sheetNum=getSheetNum()+1;
        currentSheet=createSheet("sheet"+sheetNum,prop);
    }
    public boolean writeRow(Map<String,Object> valueMap) throws IOException{
        if(totalRow>0 && totalRow % sheetMaxRows==0){
            log.debug(" finish sheet "+currentSheet.getIndex());
            currentSheet.finish();
            beginWrite();
        }
        currentSheet.writeRow(valueMap);
        totalRow++;
        return true;
    }

}
