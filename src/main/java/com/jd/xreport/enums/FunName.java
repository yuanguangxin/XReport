package com.jd.xreport.enums;

/**
 * Created by yuanguangxin on 2017/8/28.
 */
public enum FunName {
    AVG("avg"), SUM("sum"), COUNT("count"), SELECT("select"),GROUP("group");

    private String funName;

    FunName(String name){
        this.funName = name;
    }

    public String getFunName() {
        return funName;
    }
}
