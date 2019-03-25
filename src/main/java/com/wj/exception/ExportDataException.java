package com.wj.exception;

public class ExportDataException extends RuntimeException {

    private String errCode;

    private String errMsg;

    public ExportDataException(String errMsg) {
        super(errMsg);
        this.errMsg = errMsg;
    }

    public ExportDataException(String errMsg, Throwable e) {
        super(errMsg, e);
        this.errMsg = errMsg;
    }

    public ExportDataException(String errCode, String errMsg, Throwable e) {
        super(errMsg, e);
        this.errCode = errCode;
        this.errMsg = errMsg;
    }
}
