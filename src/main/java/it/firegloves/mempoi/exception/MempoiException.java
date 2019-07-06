package it.firegloves.mempoi.exception;

import java.util.concurrent.CompletionException;

public class MempoiException extends CompletionException {

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
}
