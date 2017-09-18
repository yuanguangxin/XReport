package com.jd.xreport.exception;

/**
 * Created by yuanguangxin on 2017/8/15.
 */
public class ExpException extends Exception {
    public ExpException() {
        super("Expression format error");
    }

    public ExpException(String mes) {
        super(mes);
    }
}
