package it.firegloves.exception;

public class MempoiExceptionHandler implements Thread.UncaughtExceptionHandler {

    @Override
    public void uncaughtException(Thread t, Throwable e) {

        e.printStackTrace();
    }
}
