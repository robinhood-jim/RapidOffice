package com.robin.rapidoffice.word.elements;

import com.robin.rapidoffice.meta.RelationShip;
import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.word.Document;
import lombok.Data;
import org.springframework.util.CollectionUtils;

import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

@Data
public class XRunElement implements IBodyElement{
    private String id;
    private List<PictureData> pictureDatas=new ArrayList<>();
    private String content;
    private List<TableElement> tables=new ArrayList<>();

    @Override
    public BodyType getBodyType() {
        return BodyType.RUN;
    }

    @Override
    public RelationShip getRelation() {
        return null;
    }

    @Override
    public BodyElementType getType() {
        return BodyElementType.RUN;
    }


    @Override
    public String toString(){
        StringBuilder builder=new StringBuilder();
        if(content!=null){
            builder.append("run"+id+"="+content+"\n");
        }
        if(!CollectionUtils.isEmpty(pictureDatas)){
            builder.append("pic=").append(pictureDatas.stream().map(PictureData::getRId).collect(Collectors.joining()));
            builder.append("path=").append(pictureDatas.stream().map(PictureData::getPath).collect(Collectors.joining()));
        }
        return builder.toString();
    }
}
