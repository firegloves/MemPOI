package it.firegloves.mempoi.exception;

public class MempoiRuntimeException extends RuntimeException {

    public MempoiRuntimeException() {
    }

    public MempoiRuntimeException(String message) {
        super(message);
    }

    public MempoiRuntimeException(String message, Throwable cause) {
        super(message, cause);
    }

    public MempoiRuntimeException(Throwable cause) {
        super(cause);
    }

    public MempoiRuntimeException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}
