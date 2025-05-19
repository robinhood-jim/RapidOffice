package com.robin.rapidoffice.excel.elements;

import com.robin.rapidoffice.elements.CellAddress;
import com.robin.rapidoffice.elements.CellType;
import com.robin.rapidoffice.excel.WorkBook;
import com.robin.rapidoffice.meta.Formula;
import com.robin.rapidoffice.meta.ShardingString;
import com.robin.rapidoffice.excel.utils.CellProcessor;
import com.robin.rapidoffice.excel.utils.CellUtils;
import com.robin.rapidoffice.writer.XMLWriter;
import lombok.Getter;

import java.io.IOException;

@Getter
public class Cell {
    private WorkBook workBook;
    private CellType type;
    private Object value;
    private CellAddress address;
    private String formula;
    private String rawValue;
    private String dataFormatId;
    private String dataFormatString;
    private int style;
    public Cell(WorkBook workbook, CellType type, Object value, CellAddress address, String formula, String rawValue) {
        this(workbook, type, value, address, formula, rawValue, null, null);
    }
    public Cell(WorkBook workBook,CellType type,CellAddress address){
        this.workBook=workBook;
        this.type=type;
        this.address=address;
    }
    public Cell(WorkBook workBook, CellProcessor processor,CellAddress address){
        this(workBook,processor.getCellType(),processor.getValue(),address,processor.getFormula(),processor.getRawValue());
    }

    public Cell(WorkBook workbook, CellType type, Object value, CellAddress address, String formula, String rawValue,
         String dataFormatId, String dataFormatString) {
        this.workBook = workbook;
        this.type = type;
        this.value = value;
        this.address = address;
        this.formula = formula;
        this.rawValue = rawValue;
        this.dataFormatId = dataFormatId;
        this.dataFormatString = dataFormatString;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public void setStyle(int style) {
        this.style = style;
    }

    public void write(XMLWriter w, int r, int c)  throws IOException{
        if (value != null || style != 0) {
            w.append("<c r=\"").append(CellUtils.colToString(c)).append(r).append("\"");
            if (style != 0) {
                w.append(" s=\"").append(style).append("\"");
            }
            if (value != null && !(value instanceof Formula)) {
                w.append(" t=\"").append(getCellType(value)).append("\"");
            }
            w.append(">");
            if (value instanceof Formula) {
                w.append("<f>").append(((Formula) value).getExpression()).append("</f>");
            } else if (value instanceof String) {
                w.append("<is><t>").appendEscaped((String) value);
                w.append("</t></is>");
            } else if (value != null) {
                w.append("<v>");
                if (value instanceof ShardingString) {
                    w.append(((ShardingString) value).getIndex());
                } else if (value instanceof Integer) {
                    w.append((int) value);
                } else if (value instanceof Long) {
                    w.append((long) value);
                } else if (value instanceof Double) {
                    w.append((double) value);
                } else if (value instanceof Boolean) {
                    w.append((Boolean) value ? '1' : '0');
                } else {
                    w.append(value.toString());
                }
                w.append("</v>");
            }
            w.append("</c>");
        }
    }
    static String getCellType(Object value) {
        if (value instanceof ShardingString) {
            return "s";
        } else if (value instanceof Boolean) {
            return "b";
        } else if (value instanceof String) {
            return "inlineStr";
        } else {
            return "n";
        }
    }

}
