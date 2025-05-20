package com.robin.rapidoffice.word.elements;

import lombok.Getter;

@Getter
public class LsdException {
    private String name;
    private String qFormat;
    private String semiHidden;
    private String unhideWhenUsed;
    private String uiPriority;
    private LsdException(){

    }
    public static class Builder{
        private static LsdException e=new LsdException();
        private Builder(){

        }
        public static Builder newBuilder(){
            return new Builder();
        }
        public Builder name(String name){
            e.name=name;
            return this;
        }
        public Builder qFormat(String qFormat){
            e.qFormat=qFormat;
            return this;
        }
        public Builder semiHidden(String semiHidden){
            e.semiHidden=semiHidden;
            return this;
        }
        public Builder unhideWhenUsed(String unhideWhenUsed){
            e.unhideWhenUsed=unhideWhenUsed;
            return this;
        }
        public Builder uiPriority(String uiPriority){
            e.uiPriority=uiPriority;
            return this;
        }
        public LsdException build(){
            return e;
        }
    }

}
