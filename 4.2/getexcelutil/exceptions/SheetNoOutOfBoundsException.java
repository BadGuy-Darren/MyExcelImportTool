package com.foxconn.indint.utils.getexcelutil.exceptions;

// sheet下标溢出异常
public class SheetNoOutOfBoundsException extends RuntimeException {

    public SheetNoOutOfBoundsException(String message) {
        super(message);
    }

}
