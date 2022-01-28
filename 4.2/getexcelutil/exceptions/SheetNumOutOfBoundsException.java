package com.foxconn.indint.utils.getexcelutil.exceptions;

// 请求的工作表数大于文件中包含的工作表数
public class SheetNumOutOfBoundsException extends RuntimeException {

    public SheetNumOutOfBoundsException(String message) {
        super(message);
    }

}
