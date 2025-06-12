package com.robin.rapidoffice.excel;

import com.robin.comm.util.xls.ExcelSheetProp;
import lombok.extern.slf4j.Slf4j;

import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.function.Consumer;

@Slf4j
public class SingleWorkBook extends WorkBook {
    private ExcelSheetProp prop;
    private WorkSheet currentSheet;
    private int sheetMaxRows;
    int totalRow=0;
    int thresholdSize =8*1024;

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
        if(prop.getMaxRows()>0) {
            this.sheetMaxRows = prop.getMaxRows();
        }
        if(prop.getMaxSheetSize()>0){
            this.maxSheetSize=prop.getMaxSheetSize();
        }
    }
    public SingleWorkBook(File path, int bufferSize, ExcelSheetProp prop,int threshold) {
        this(path, bufferSize,prop);
        if(threshold>0) {
            this.thresholdSize = threshold;
        }
    }


    public SingleWorkBook(OutputStream outputStream, ExcelSheetProp prop) {
        super(outputStream);
        this.prop=prop;
        if(prop.getMaxRows()>0){
            sheetMaxRows=prop.getMaxRows();
        }
        if(prop.getMaxSheetSize()>0){
            this.maxSheetSize=prop.getMaxSheetSize();
        }
    }

    public void beginWrite() throws IOException{
        int sheetNum=getSheetNum()+1;
        currentSheet=createSheet("sheet"+sheetNum,prop);
    }
    public void beginWrite(Consumer<WorkSheet> consumer) throws IOException{
        int sheetNum=getSheetNum()+1;
        currentSheet=createSheet("sheet"+sheetNum,prop,consumer);
    }
    public boolean writeRow(Map<String,Object> valueMap) throws IOException{
        if(totalRow>0 && (totalRow % sheetMaxRows==0 || (maxSheetSize>0 && sheetWriterMap.get(currentSheet.getIndex()).shouldClose(maxSheetSize, thresholdSize)))){
            log.debug(" finish sheet "+currentSheet.getIndex());
            currentSheet.finish();
            beginWrite();
        }
        currentSheet.writeRow(valueMap);
        totalRow++;
        return true;
    }

}
