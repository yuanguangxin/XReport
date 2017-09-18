package com.jd.xreport.models;

import com.jd.xreport.exception.DSException;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuanguangxin on 2017/8/15.
 *
 * 数据集容器类
 */
public class DSContainer {
    private static List<DataSet> dataSets = new ArrayList<DataSet>();

    public static DataSet getDataSet(String name) {
        for (int i = 0; i < dataSets.size(); i++) {
            if (name.equals(dataSets.get(i).getName())) {
                return dataSets.get(i);
            }
        }
        return null;
    }

    public static void addDataSet(DataSet dataSet) throws DSException {
        for (int i = 0; i < dataSets.size(); i++) {
            if (dataSet.getName().equals(dataSets.get(i).getName())) {
                throw new DSException("DataSource '" + dataSet.getName() + "' is already exists");
            }
        }
        dataSets.add(dataSet);
    }

    public static void removeDataSet(String name) {
        for (int i = 0; i < dataSets.size(); i++) {
            if (name.equals(dataSets.get(i).getName())) {
                dataSets.remove(i);
                break;
            }
        }
    }
}
