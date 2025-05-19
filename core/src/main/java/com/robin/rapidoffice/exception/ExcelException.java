package com.robin.rapidoffice.exception;

public class ExcelException extends RuntimeException{
    public ExcelException(Exception ex){
        super(ex);
    }
    public ExcelException(String message){
        super(message);
    }
}
