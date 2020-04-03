package com.kovizone.poi.ooxml.plus.exception;

/**
 * Poiplus异常
 *
 * @author KoviChen
 */
public class PoiOoxmlPlusException extends Exception {

    public PoiOoxmlPlusException() {
        super();
    }

    public PoiOoxmlPlusException(String message) {
        super(message);
    }

    public PoiOoxmlPlusException(String message, Throwable cause) {
        super(message, cause);
    }

    public PoiOoxmlPlusException(Throwable cause) {
        super(cause);
    }

    protected PoiOoxmlPlusException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
