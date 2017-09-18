package com.jd.xreport.utils;

import com.jd.xreport.models.ExcelData;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuanguangxin on 2017/9/15.
 */
public class ListUtil {
    public static void joinLeft(List<List<ExcelData>> left, List<List<ExcelData>> right) {
        int n = right.size();
        int m = 0;
        if (right.size() > 0) m = right.get(0).size();
        for (int i = 0; i < left.size(); i++) {
            if (i >= n) {
                List<ExcelData> temp = new ArrayList<>();
                for (int k = 0; k < m; k++) {
                    temp.add(null);
                }
                left.get(i).addAll(temp);
                continue;
            }
            left.get(i).addAll(right.get(i));
        }
    }

    public static void joinTop(List<List<ExcelData>> top, List<List<ExcelData>> below) {
        top.addAll(below);
    }
}
