package com.jd.xreport.models;

import com.jd.xreport.enums.FunName;
import com.jd.xreport.exception.DSException;
import com.jd.xreport.exception.ExpException;
import com.jd.xreport.utils.DataUtil;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuanguangxin on 2017/8/15.
 *
 * 表达式类
 */
@Data
public class DSExpression {
    private DataSet dataSet;//表达式所属数据集
    private String exp;//表达式
    private List<String> result;//表达式解析后的结果集

    public DSExpression(String exp) throws DSException, ExpException {
        this.exp = exp.substring(exp.indexOf(".") + 1, exp.length());
        this.dataSet = DSContainer.getDataSet(exp.substring(1, exp.indexOf(".")));
        if (dataSet == null)
            throw new DSException("dataSet '" + exp.substring(1, exp.length()).split("\\.")[0] + "' not found");
    }

    public void analysis() throws ExpException {
        String[] args;
        FunName funName;
        try {
            funName = FunName.valueOf(exp.split("\\(")[0].toUpperCase());
            int index = exp.indexOf("(");
            args = exp.substring(index + 1, exp.length() - 1).split(",");
        } catch (Exception e) {
            throw new ExpException();
        }
        switch (funName) {
            case AVG:
                if (args.length == 1) {
                    result = DataUtil.avg(dataSet.getData(), args[0]);
                } else if (args.length == 2) {
                    result = DataUtil.avg(dataSet.getData(), args[0], args[1]);
                }
                break;
            case SUM:
                if (args.length == 1) {
                    result = DataUtil.sum(dataSet.getData(), args[0]);
                } else if (args.length == 2) {
                    result = DataUtil.sum(dataSet.getData(), args[0], args[1]);
                }
                break;
            case COUNT:
                if (args[0].equals("")) {
                    List<String> list = new ArrayList<>();
                    list.add(String.valueOf(dataSet.getData().size()));
                    result = list;
                } else {
                    result = DataUtil.count(dataSet.getData(), args[0]);
                }
                break;
            case SELECT:
                if (args.length == 1) {
                    result = DataUtil.select(dataSet.getData(), args[0]);
                } else if (args.length == 2) {
                    result = DataUtil.select(dataSet.getData(), args[0], args[1]);
                }
                break;
            case GROUP:
                if (args.length == 1) {
                    result = DataUtil.group(dataSet.getData(), args[0]);
                } else if (args.length == 2) {
                    if (args[1].contains(";")) {
                        List<String> list = new ArrayList<>();
                        String[] s = args[1].split(";");
                        for (int i = 0; i < s.length; i++) {
                            list.add(s[i]);
                        }
                        result = DataUtil.group(dataSet.getData(), args[0], list);
                    } else {
                        result = DataUtil.group(dataSet.getData(), args[0], args[1]);
                    }
                } else if (args.length == 3) {
                    List<String> list = new ArrayList<>();
                    String[] s = args[1].split(";");
                    for (int i = 0; i < s.length; i++) {
                        list.add(s[i]);
                    }
                    result = DataUtil.group(dataSet.getData(), args[0], list, args[2]);
                }
                break;
        }
    }
}
