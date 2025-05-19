package com.robin.rapidoffice.excel;

import cn.hutool.core.util.NumberUtil;

import com.robin.comm.util.xls.ExcelColumnProp;
import com.robin.comm.util.xls.ExcelSheetProp;
import com.robin.core.base.util.Const;
import com.robin.core.fileaccess.util.ByteBufferOutputStream;
import com.robin.rapidoffice.elements.*;
import com.robin.rapidoffice.excel.elements.Cell;
import com.robin.rapidoffice.excel.utils.CellUtils;
import com.robin.rapidoffice.excel.utils.DateUtils;
import com.robin.rapidoffice.exception.ExcelException;
import com.robin.rapidoffice.meta.Fill;
import com.robin.rapidoffice.meta.Font;
import com.robin.rapidoffice.meta.Formula;
import com.robin.rapidoffice.meta.ShardingString;

import com.robin.rapidoffice.utils.OPCPackage;
import com.robin.rapidoffice.writer.XMLWriter;
import org.springframework.util.Assert;
import org.springframework.util.ObjectUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.function.Consumer;

public class WorkSheet {
    private String id;
    private int index;
    private String name;
    private String sheetId;
    private SheetVisibility visibility;
    private WorkBook workBook;
    boolean finished = false;
    private final Set<Integer> hiddenRows = new HashSet<>();

    private final Set<Integer> hiddenColumns = new HashSet<>();
    private final Map<Integer, Double> colWidths = new HashMap<>();

    private final Map<Integer, Column> colStyles = new HashMap<>();

    private Boolean fitToPage = false;
    private Boolean autoPageBreaks = false;
    private Cell[] currentCells;
    private int currentRowNum=1;
    private List<Integer> styles=new ArrayList<>();
    ExcelSheetProp prop;

    Font defaultFont;
    Fill defaultFill;
    Border defaultBorder;
    Alignment defaultAlignment;
    boolean ifHidden;
    public static final int MAX_ROWS = 1_048_576;

    public WorkSheet(WorkBook workBook,ExcelSheetProp prop,int index,String id,String sheetId,String name,SheetVisibility visibility){
        this(workBook,index,id,sheetId,name,visibility);
        this.prop=prop;
    }
    public WorkSheet(WorkBook workBook,int index,String id,String sheetId,String name,SheetVisibility visibility){
        this.index=index;
        this.id=id;
        this.name=name;
        this.visibility=visibility;
        this.sheetId=sheetId;
        this.workBook=workBook;
    }
    public WorkSheet(WorkBook workBook,String name){
        this.workBook=workBook;
        this.name=name;
    }

    public void writeHeader(XMLWriter w) throws IOException {

        w.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        w.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        w.append("<sheetPr filterMode=\"" + "false" + "\">");

        w.append("<pageSetUpPr fitToPage=\"" + fitToPage + "\" " + "autoPageBreaks=\"" + autoPageBreaks + "\"/></sheetPr>");
        w.append("<dimension ref=\"A1\"/>");
        w.append("<sheetViews><sheetView workbookViewId=\"0\">");
        w.append("</sheetView>");
        w.append("</sheetViews><sheetFormatPr defaultRowHeight=\"15.0\"/>");
        w.append(getColumnDefine());
        w.append("<sheetData>");
    }
    public Cell getCell(int col, ExcelSheetProp prop){
        if(currentCells ==null){
            currentCells =new Cell[prop.getColumnPropList().size()];
            for(int i=0;i<prop.getColumnPropList().size();i++){
                Cell cell=new Cell(workBook, OPCPackage.parseCellType(prop.getColumnPropList().get(i)),new CellAddress(currentRowNum,i+1));
                cell.setStyle(styles.get(i));
                if(prop.getColumnPropList().get(i).getColumnType().equals(Const.META_TYPE_FORMULA)){
                    cell.setValue(new Formula(prop.getColumnPropList().get(i).getFormula()));
                }
                currentCells[i]=cell;
            }
        }
        if(col<=prop.getColumnPropList().size()) {
            return currentCells[col];
        }else{
            throw new ExcelException("row over excel define columns");
        }
    }
    public void writeRow(Map<String,Object> valueMap) throws IOException{
        XMLWriter w=workBook.sheetWriterMap.get(getIndex());
        if(currentRowNum==MAX_ROWS-1){
            throw new ExcelException("sheet over row limit!");
        }
        if(!ObjectUtils.isEmpty(w)){
            for(int i=0;i<prop.getColumnPropList().size();i++){
                ExcelColumnProp columnProp=prop.getColumnPropList().get(i);
                if(!ObjectUtils.isEmpty(valueMap.get(columnProp.getColumnCode()))){
                    Cell cell=getCell(i,prop);
                    setValue(cell,columnProp,valueMap.get(columnProp.getColumnCode()));
                }else{
                    if(getCell(i,prop).getType()!=CellType.FORMULA) {
                        getCell(i, prop).setValue(null);
                    }else{
                        Formula formula=(Formula)getCell(i,prop).getValue();
                        formula.setExpression(CellUtils.returnFormulaWithPos(columnProp.getFormula(),currentRowNum));
                    }
                }
            }
            writeRow(w,ifHidden,(byte)0,null);
            currentRowNum++;
        }
    }
    void writeTitle(XMLWriter w, ExcelSheetProp prop) throws IOException{
        w.append("<row r=\"").append(currentRowNum).append("\" s=\"1\" >");
        for(int i=0;i<prop.getColumnPropList().size();i++) {
            ShardingString s1=workBook.addShardingString(prop.getColumnPropList().get(i).getColumnName());
            w.append("<c r=\"").append(CellUtils.colToString(i)).append(currentRowNum).append("\" t=\"s\" >")
                    .append("<v>").append(s1.getIndex()).append("</v></c>");
        }
        w.append("</row>");
    }
    void writeRow(XMLWriter w,boolean isHidden,byte groupLevel,
                     Double rowHeight) throws IOException{

        w.append("<row r=\"").append(currentRowNum).append("\"");
        if (isHidden) {
            w.append(" hidden=\"true\"");
        }
        if(rowHeight != null) {
            w.append(" ht=\"")
                    .append(rowHeight)
                    .append("\"")
                    .append(" customHeight=\"1\"");
        }
        if (groupLevel!=0){
            w.append(" outlineLevel=\"")
                    .append(groupLevel)
                    .append("\"");
        }
        w.append(">");
        for (int c = 0; c < currentCells.length; ++c) {
            if (currentCells[c] != null) {
                currentCells[c].write(w, currentRowNum, c);
            }
        }
        w.append("</row>");

    }
    void setValue(Cell cell,ExcelColumnProp prop,Object value){
        switch (prop.getColumnType()){
            case Const.META_TYPE_DECIMAL:
            case Const.META_TYPE_FLOAT:
            case Const.META_TYPE_DOUBLE:
                if(NumberUtil.isNumber(value.toString())) {
                    cell.setValue(Double.valueOf(value.toString()));
                }
                break;
            case Const.META_TYPE_BIGINT:
            case Const.META_TYPE_INTEGER:
                if(NumberUtil.isNumber(value.toString())){
                    int pos=value.toString().indexOf(".");
                    if(pos!=-1){
                        cell.setValue(Integer.valueOf(value.toString().substring(0,pos)));
                    }else{
                        cell.setValue(Integer.valueOf(value.toString()));
                    }
                }
                break;
            case Const.META_TYPE_TIMESTAMP:
            case Const.META_TYPE_DATE:
                if(LocalDateTime.class.isAssignableFrom(value.getClass())) {
                    cell.setValue(DateUtils.convertDate((LocalDateTime) value));
                }else if(Date.class.isAssignableFrom(value.getClass())){
                    cell.setValue(DateUtils.convertDate((Date)value));
                }else if(LocalDate.class.isAssignableFrom(value.getClass())){
                    cell.setValue(DateUtils.convertDate((LocalDate) value));
                }else if(String.class.isAssignableFrom(value.getClass())){
                    cell.setValue(DateUtils.convertDate(new Timestamp(Long.valueOf(value.toString()))));
                }
                break;
            case Const.META_TYPE_FORMULA:
                Formula formula=(Formula) cell.getValue();
                formula.setExpression(CellUtils.returnFormulaWithPos(prop.getFormula(),currentRowNum));
                break;
            default:
                cell.setValue(value);
        }
    }
    private void mergeCell(CellAddress begin,CellAddress end,Object value){

    }
    void setDefaultStyles(Consumer<WorkBook> consumer){
        Assert.notNull(prop,"");
        styles=new ArrayList<>(prop.getColumnPropList().size());
        if(consumer!=null){
            consumer.accept(workBook);
        }else {
            defaultFont = getDefaultFont();
            defaultFill=getDefaultFill();
            defaultBorder=getDefaultBorder();
            defaultAlignment=getDefaultAlignment();
            prop.getColumnPropList().forEach(f -> styles.add(workBook.holder.mergeCellStyle(0, getFormatWithType(f), defaultFont,defaultFill,defaultBorder, defaultAlignment)));
        }
    }
    Font getDefaultFont(){
        Font font=new Font(false,false,false,CellUtils.getDefaultFontName(),BigDecimal.valueOf(12.0),null,false);
        return font;
    }
    Fill getDefaultFill(){
        return Fill.BLACK;
    }
    Border getDefaultBorder(){
        return Border.BLACK;
    }
    Alignment getDefaultAlignment(){
        return new Alignment("center","center",false,0,0);
    }
    String getColumnDefine() throws IOException{
        return "<cols><col min=\"1\" max=\"1\" customWidth=\"true\"></col></cols>";
    }
    public String getFormatWithType(ExcelColumnProp columnProp){
        Assert.notNull(columnProp,"");
        String numFmtStr;
        switch (columnProp.getColumnType()){
            case Const.META_TYPE_BIGINT:
            case Const.META_TYPE_INTEGER:
                numFmtStr=ObjectUtils.isEmpty(columnProp.getFormat())?OPCPackage.IMPLICIT_NUM_FMTS.get("1"):columnProp.getFormat();
                break;
            case Const.META_TYPE_DECIMAL:
            case Const.META_TYPE_NUMERIC:
            case Const.META_TYPE_DOUBLE:
            case Const.META_TYPE_FLOAT:
                numFmtStr=ObjectUtils.isEmpty(columnProp.getFormat())?OPCPackage.IMPLICIT_NUM_FMTS.get("2"):columnProp.getFormat();
                break;
            case Const.META_TYPE_BOOLEAN:
                numFmtStr=ObjectUtils.isEmpty(columnProp.getFormat())?OPCPackage.IMPLICIT_NUM_FMTS.get("1"):columnProp.getFormat();
                break;
            case Const.META_TYPE_DATE:
            case Const.META_TYPE_TIMESTAMP:
                numFmtStr=ObjectUtils.isEmpty(columnProp.getFormat())?"yyyy-MM-dd hh:mm:ss":columnProp.getFormat();
                break;
            case Const.META_TYPE_FORMULA:
                numFmtStr=OPCPackage.IMPLICIT_NUM_FMTS.get("2");
                break;
            default:
                numFmtStr=OPCPackage.IMPLICIT_NUM_FMTS.get("1");
        }
        return numFmtStr;
    }
    public void finish() throws IOException{
        if(finished){
            return;
        }
        XMLWriter w=workBook.sheetWriterMap.get(getIndex());

        w.append("</sheetData>");
        w.append("</worksheet>");

        w.flush();
        OutputStream outputStream=workBook.sheetTmpStreamMap.get(getIndex());
        if(!ByteBufferOutputStream.class.isAssignableFrom(outputStream.getClass())) {
            outputStream.close();
        }
        finished=true;
    }

    public void setStyle(int pos ,Font font,String format,Fill fill){
        Assert.isTrue(pos>0,"");
        if(styles.size()>pos){
            styles.set(pos-1,workBook.holder.mergeCellStyle(0, format, font,fill,defaultBorder, defaultAlignment));
        }
    }

    public String getId() {
        return id;
    }

    public int getIndex() {
        return index;
    }

    public String getName() {
        return name;
    }

    public SheetVisibility getVisibility() {
        return visibility;
    }
}
