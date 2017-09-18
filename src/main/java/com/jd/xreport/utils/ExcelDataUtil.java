package com.jd.xreport.utils;

import com.jd.xreport.models.ExcelData;
import com.jd.xreport.models.Point;

import java.util.List;

/**
 * Created by yuanguangxin on 2017/9/12.
 */
public class ExcelDataUtil {
    public static Point getHorPoint(List<List<ExcelData>> lists) {
        Point point = new Point();
        for (int i = 0; i < lists.size(); i++) {
            for (int j = 0; j < lists.get(i).size(); j++) {
                if (lists.get(i).get(j).getValue() != null && lists.get(i).get(j).getValue().contains("#")) {
                    point.setX(i);
                    point.setY(j);
                    return point;
                }
            }
        }
        point.setX(-1);
        point.setY(-1);
        return point;
    }

    public static Point getVerPoint(List<List<ExcelData>> lists) {
        Point point = new Point();
        for (int i = 0; i < lists.size(); i++) {
            for (int j = 0; j < lists.get(i).size(); j++) {
                if (lists.get(i).get(j).getValue() != null && lists.get(i).get(j).getValue().contains("$")) {
                    point.setX(i);
                    point.setY(j);
                    return point;
                }
            }
        }
        point.setX(-1);
        point.setY(-1);
        return point;
    }
}
