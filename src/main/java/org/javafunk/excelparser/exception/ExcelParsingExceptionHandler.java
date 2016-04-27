package org.javafunk.excelparser.exception;

/**
 * Created by Victor.Ikoro on 4/27/2016.
 */
public class ExcelParsingExceptionHandler implements IExceptionHandler {

    private ExcelParsingException exception;

    public ExcelParsingExceptionHandler()
    {
        this(new ExcelParsingException(""));
    }
    public ExcelParsingExceptionHandler(ExcelParsingException ex)
    {
        this.exception = ex;
    }

    @Override
    public void handle() {
        throw  this.exception;
    }

    @Override
    public void setException(Exception ex) {
        this.exception = (ExcelParsingException) ex;
    }


}
