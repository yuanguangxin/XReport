package com.jd.xreport.models;

import com.jd.xreport.c3p0.DBUtil;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * Created by yuanguangxin on 2017/8/15.
 *
 * 数据集类
 */
@Data
public class DataSet {
    private String sql;//数据集对应sql
    private String name;//数据集名称
    private List<Map<String, String>> data;//数据集

    public DataSet(String name, String sql) {
        this.name = name;
        this.sql = sql;
        this.data = DBUtil.executeQuery(sql);
    }
}
