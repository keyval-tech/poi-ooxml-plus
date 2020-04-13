package com.kovizone.poi.ooxml.plus.exception;

public class ReflexException extends Exception {

    public ReflexException() {
        super();
    }

    public ReflexException(String message) {
        super(message);
    }

    public ReflexException(String message, Throwable cause) {
        super(message, cause);
    }

    public ReflexException(Throwable cause) {
        super(cause);
    }

    protected ReflexException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
