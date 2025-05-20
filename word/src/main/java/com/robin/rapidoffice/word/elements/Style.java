package com.robin.rapidoffice.word.elements;

import lombok.Data;

@Data
public class Style {
    private String styleId;
    private String type;
    private String defaultVal;
    private String baseOn;
    private String name;
    private String qFormat;
    private String uiPriority;
    private PrDefaultRpr rpr;

}
