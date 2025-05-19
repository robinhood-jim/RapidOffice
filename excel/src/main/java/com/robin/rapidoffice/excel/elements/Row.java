package com.robin.rapidoffice.excel.elements;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class Row {
    private List<Cell> cells;
    public Row(List<Cell> cells){
        this.cells=cells;
    }
    private StringBuilder builder=new StringBuilder();
    @Override
    public String toString() {
        if(builder.length()>0){
            builder.delete(0,builder.length());
        }
        builder.append("{");
        for(int i=0;i<cells.size();i++){
            builder.append(i).append(":").append(cells.get(i).getValue());
            if(i<cells.size()-1){
                builder.append(",");
            }
        }
        builder.append("}");
        return builder.toString();
    }
}
