package com.simpleWorkbook.exception;

public class FileTypeNotSupportException extends RuntimeException {

    public FileTypeNotSupportException() {
        super();
    }

    public FileTypeNotSupportException(String message) {
        super(message);
    }

    public FileTypeNotSupportException(String message, Throwable cause) {
        super(message, cause);
    }

    public FileTypeNotSupportException(Throwable cause) {
        super(cause);
    }
}
