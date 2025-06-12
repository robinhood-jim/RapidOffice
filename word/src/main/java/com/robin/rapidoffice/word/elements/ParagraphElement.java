package com.robin.rapidoffice.word.elements;

import com.robin.rapidoffice.meta.RelationShip;
import lombok.Getter;
import org.springframework.util.CollectionUtils;

import java.util.List;

@Getter
public class ParagraphElement implements IBodyElement {
    private String id;
    private String rsidR;
    private String rsidRDefault;
    private XRunElement element;
    private List<XRunElement> elements;
    private String runId;

    public ParagraphElement(String id,String rsidR,String rsidRDefault, String runId, List<XRunElement> elements){
        this.id=id;
        this.runId=runId;
        this.elements=elements;
        this.rsidRDefault=rsidRDefault;
        this.rsidR=rsidR;
    }

    @Override
    public BodyType getBodyType() {
        return BodyType.DOCUMENT;
    }

    @Override
    public RelationShip getRelation() {
        return null;
    }

    @Override
    public BodyElementType getType() {
        return BodyElementType.PARAGRAPH;
    }


    @Override
    public String toString() {
        StringBuilder builder=new StringBuilder("Paragraph \n");
        if(!CollectionUtils.isEmpty(elements)){
            elements.stream().forEach(f->builder.append(f.toString()));
        }
        builder.append("Paragraph end!");
        return builder.toString();
    }
}
