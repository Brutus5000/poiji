package com.poiji.exception;

/**
 * Thrown if the model class to be used for parsing an Excel sheet violates any constraints
 */
public class InvalidModelException extends PoijiException {
    public InvalidModelException(String message) {
        super(message);
    }
}
