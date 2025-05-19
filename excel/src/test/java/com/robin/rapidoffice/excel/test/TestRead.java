package com.robin.rapidoffice.excel.test;

import com.robin.comm.util.xls.ExcelColumnProp;
import com.robin.comm.util.xls.ExcelSheetProp;
import com.robin.core.base.util.Const;
import com.robin.core.base.util.StringUtils;
import com.robin.rapidoffice.excel.elements.Row;
import com.robin.rapidoffice.excel.SingleWorkBook;
import com.robin.rapidoffice.excel.WorkBook;
import com.robin.rapidoffice.excel.WorkSheet;
import org.junit.Test;

import javax.xml.stream.XMLStreamException;
import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Random;
import java.util.stream.Stream;

public class TestRead {
    @Test
    public void testRead() throws IOException{
        //new BufferedReader(new InputStreamReader(System.in)).readLine();
        ExcelSheetProp.Builder builder = ExcelSheetProp.Builder.newBuilder();
        builder.addColumnProp(new ExcelColumnProp("name", "name", Const.META_TYPE_STRING, false))
        .addColumnProp("time", "time", Const.META_TYPE_TIMESTAMP, false)
        .addColumnProp("intcol", "intcol", Const.META_TYPE_INTEGER, false)
        .addColumnProp("dval", "dval", Const.META_TYPE_DOUBLE, false)
        .addColumnProp("dval2", "dval2", Const.META_TYPE_DOUBLE, false)
        .addColumnProp(new ExcelColumnProp("diff", "diff", Const.META_TYPE_FORMULA, "(D{P}-E{P})/C{P}")).setStreamMode();
        try(WorkBook workBook=new WorkBook(new File("D:/test111.xlsx"))){
            int sheetNum= workBook.getSheetNum();
            ExcelSheetProp prop=builder.build();
            for(int i=0;i<sheetNum;i++){
                WorkSheet sheet=workBook.getSheet(i).get();
                Stream<Row> stream= workBook.openStream(sheet,prop);
                Iterator<Row> iter=stream.iterator();
                int pos=0;
                while(iter.hasNext()){
                    iter.next();
                    //System.out.println(iter.next());
                    pos++;
                }
                System.out.println("sheet "+sheet.getName()+" rows "+pos);
            }

        }catch (IOException| XMLStreamException ex){
            ex.printStackTrace();
        }
        //new BufferedReader(new InputStreamReader(System.in)).readLine();
    }
    @Test
    public void testWrite() throws IOException {
        //new BufferedReader(new InputStreamReader(System.in)).readLine();
        ExcelSheetProp.Builder builder = ExcelSheetProp.Builder.newBuilder();
        builder.addColumnProp(new ExcelColumnProp("name", "name", Const.META_TYPE_STRING, false))
                .addColumnProp("time", "time", Const.META_TYPE_TIMESTAMP, false)
                .addColumnProp("intcol", "intcol", Const.META_TYPE_INTEGER, false)
                .addColumnProp("dval", "dval", Const.META_TYPE_DOUBLE, false)
                .addColumnProp("dval2", "dval2", Const.META_TYPE_DOUBLE, false)
                .addColumnProp(new ExcelColumnProp("diff", "diff", Const.META_TYPE_FORMULA, "(D{P}-E{P})/C{P}")).setStreamMode();
        Random random = new Random(12312321321312L);
        Map<String,Object> cachedMap=new HashMap<>();
        try(SingleWorkBook workBook=new SingleWorkBook(new File("d:/test111.xlsx"),0,builder.build())){
            Long startTs = System.currentTimeMillis() - 3600 * 24 * 1000;
            workBook.beginWrite();
            for(int j=0;j<1200000;j++){
                cachedMap.put("name", StringUtils.generateRandomChar(12));
                cachedMap.put("time", String.valueOf(startTs + j * 1000));
                cachedMap.put("intcol", String.valueOf(random.nextInt(1000)));
                cachedMap.put("dval", String.valueOf(random.nextDouble() * 1000));
                cachedMap.put("dval2", String.valueOf(random.nextDouble() * 500));
                workBook.writeRow(cachedMap);
            }
        }
    }
}
