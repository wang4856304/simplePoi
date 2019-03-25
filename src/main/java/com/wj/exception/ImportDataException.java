package com.wj.exception;

public class ImportDataException extends RuntimeException {

    private String errCode;

    private String errMsg;

    public ImportDataException(String errMsg) {
        super(errMsg);
        this.errMsg = errMsg;
    }

    public ImportDataException(String errMsg, Throwable e) {
        super(errMsg, e);
        this.errMsg = errMsg;
    }

    public ImportDataException(String errCode, String errMsg, Throwable e) {
        super(errMsg, e);
        this.errCode = errCode;
        this.errMsg = errMsg;
    }
}
