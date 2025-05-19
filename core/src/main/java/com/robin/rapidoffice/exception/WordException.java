package com.robin.rapidoffice.exception;

public class WordException extends RuntimeException{
    public WordException(Exception ex){
        super(ex);
    }
    public WordException(String message){
        super(message);
    }
}
