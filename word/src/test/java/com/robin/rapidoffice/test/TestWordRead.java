package com.robin.rapidoffice.test;

import com.robin.rapidoffice.word.Document;
import com.robin.rapidoffice.word.elements.IBodyElement;
import org.junit.Test;

import java.io.File;
import java.util.Iterator;

public class TestWordRead {
    @Test
    public void doRead(){
        try(Document document=new Document(new File("f:/test123.docx"))){
            Iterator<IBodyElement> iter=document.openStream().iterator();
            while(iter.hasNext()){
                System.out.println(iter.next());
            }
        }catch (Exception ex){
            ex.printStackTrace();
        }
    }
}
