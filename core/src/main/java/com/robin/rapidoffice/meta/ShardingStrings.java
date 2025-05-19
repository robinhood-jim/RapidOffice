package com.robin.rapidoffice.meta;

import com.robin.rapidoffice.reader.XMLReader;
import com.robin.rapidoffice.utils.XMLFactoryUtils;
import com.robin.rapidoffice.writer.XMLWriter;
import com.robin.rapidoffice.elements.IWriteableElements;

import javax.xml.stream.XMLStreamException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class ShardingStrings implements IWriteableElements {
    private List<ShardingString> values=new ArrayList<>();
    private InputStream inputStream;
    private ShardingStrings(InputStream inputStream){
        this.inputStream=inputStream;
    }
    public static ShardingStrings formInputStream(InputStream inputStream) throws XMLStreamException,IOException {
        ShardingStrings shardingStrings=new ShardingStrings(inputStream);
        shardingStrings.construct();
        return shardingStrings;
    }
    public ShardingStrings(){
        
    }
    private void construct() throws XMLStreamException, IOException{
        int pos=0;
        try(XMLReader reader=new XMLReader(XMLFactoryUtils.getDefaultInputFactory(),inputStream)){
            while (reader.hasNext()) {
                reader.goTo("si");
                StringBuilder sb=new StringBuilder();
                while (reader.goTo(() -> reader.isStartElement("t")
                        || reader.isStartElement("rPh")
                        || reader.isEndElement("si"))) {
                    if (reader.isStartElement("t")) {
                        sb.append(reader.getValueUntilEndElement("t"));
                    } else if (reader.isEndElement("si")) {
                        break;
                    } else if (reader.isStartElement("rPh")) {
                        reader.goTo(() -> reader.isEndElement("rPh"));
                    }
                }
                if(sb.length()>0) {
                    values.add(new ShardingString(sb.toString(),pos++));
                }
            }
        }
    }

    public List<ShardingString> getValues() {
        return values;
    }

    @Override
    public void writeOut(XMLWriter writer) throws IOException {
        writer.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
                .append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"10\" uniqueCount=\"")
        .append(values.size()+"\">\n");
        for(ShardingString value:values){
            writer.append("<si><t>").appendEscaped(value.getValue()).append("</t></si>");
        }
        writer.append("\n</sst>");
    }
}
