package org.javafunk.excelparser.exception;

/**
 * Created by Victor.Ikoro on 4/27/2016.
 */
public interface IExceptionHandler {
    public void handle();
    public  void setException(Exception ex);
}
