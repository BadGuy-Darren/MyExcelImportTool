package com.foxconn.indint.utils.getexcelutil.exceptions;

// 接收excel文件为空异常
public class EmptyExcelFileException extends RuntimeException {

    public EmptyExcelFileException(String message) {
        super(message);
    }

}
