package com.kovizone.poi.ooxml.plus.exception;

/**
 * Poiplus异常
 *
 * @author KoviChen
 */
public class ExcelWriteException extends Exception {

    public ExcelWriteException() {
        super();
    }

    public ExcelWriteException(String message) {
        super(message);
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelWriteException(Throwable cause) {
        super(cause);
    }

    protected ExcelWriteException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
