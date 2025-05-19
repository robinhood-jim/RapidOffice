package com.robin.rapidoffice.meta;


import com.robin.rapidoffice.writer.XMLWriter;
import lombok.Getter;
import lombok.Setter;

import java.io.IOException;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.Set;

@Getter
@Setter
public class Properties {
    private String title;
    private String subject;
    private String keywords;

    private String description;
    private String category;

    private String manager;
    private String company;
    private String hyperlinkBase;

    private Set<ICustomProperty> customProperties = Collections.synchronizedSet(new LinkedHashSet<>());
    interface ICustomProperty{
        void writeOut(XMLWriter writer, int pid) throws IOException;
    }
    abstract class AbstractProperty<T> implements ICustomProperty{
        protected String key;
        protected  T value;
        public AbstractProperty(String key,T value){
            this.key=key;
            this.value=value;
        }
    }


}
