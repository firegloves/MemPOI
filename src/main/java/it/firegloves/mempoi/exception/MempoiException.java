package it.firegloves.mempoi.exception;

public class MempoiException extends RuntimeException {

    public MempoiException() {
    }

    public MempoiException(String message) {
        super(message);
    }

    public MempoiException(String message, Throwable cause) {
        super(message, cause);
    }

    public MempoiException(Throwable cause) {
        super(cause);
    }

    public MempoiException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
