package com.robin.rapidoffice.exception;

public class StreamReadException extends RuntimeException{
    public StreamReadException(Exception ex){
        super(ex);
    }
    public StreamReadException(String message){
        super(message);
    }
}
