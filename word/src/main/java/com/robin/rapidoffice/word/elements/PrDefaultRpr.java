package com.robin.rapidoffice.word.elements;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
public class PrDefaultRpr {
    private Map<String,String> fonts=new HashMap<>();
    private Map<String,String> langs=new HashMap<>();
    private String sz;
    private String szCs;

}
