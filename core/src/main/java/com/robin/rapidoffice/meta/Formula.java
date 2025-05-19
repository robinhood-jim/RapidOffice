package com.robin.rapidoffice.meta;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class Formula {
    private String expression;
    public Formula(String expression){
        this.expression=expression;
    }

}
