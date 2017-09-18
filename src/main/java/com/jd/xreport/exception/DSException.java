package com.jd.xreport.exception;

/**
 * Created by yuanguangxin on 2017/8/15.
 */
public class DSException extends Exception {
    public DSException() {
        super("DataSet error");
    }

    public DSException(String mes) {
        super(mes);
    }
}
