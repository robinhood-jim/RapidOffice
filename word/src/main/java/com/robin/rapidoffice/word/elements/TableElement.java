package com.robin.rapidoffice.word.elements;

import cn.hutool.core.util.StrUtil;
import com.robin.rapidoffice.meta.RelationShip;
import lombok.Data;
import org.springframework.util.CollectionUtils;

import java.util.List;

@Data
public class TableElement implements IBodyElement {
    List<String> headers;
    List<List<String>> values;
    public TableElement(List<String> headers,List<List<String>> values){
        this.headers=headers;
        this.values=values;
    }

    @Override
    public BodyType getBodyType() {
        return BodyType.TABLECELL;
    }

    @Override
    public RelationShip getRelation() {
        return null;
    }

    @Override
    public BodyElementType getType() {
        return BodyElementType.TABLE;
    }


    @Override
    public String toString(){
        StringBuilder builder=new StringBuilder();
        builder.append("table begin \n");
        if(!CollectionUtils.isEmpty(headers)){
            builder.append("headrs\n"+ StrUtil.join(",",headers)+"\n values:\n ");
        }
        if(!CollectionUtils.isEmpty(values)){
            values.forEach(f->builder.append(StrUtil.join(",",f)+"\n"));
        }
        return builder.toString();
    }
}
