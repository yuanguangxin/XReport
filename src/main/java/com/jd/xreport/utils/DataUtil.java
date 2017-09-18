package com.jd.xreport.utils;


import com.jd.xreport.exception.ExpException;
import org.apache.commons.jexl3.JexlBuilder;
import org.apache.commons.jexl3.JexlEngine;
import org.apache.commons.jexl3.JexlExpression;

import java.util.*;

/**
 * Created by yuanguangxin on 2017/8/16.
 */
public class DataUtil {
    private static JexlEngine jexlEngine = new JexlBuilder().create();
    private static JexlExpression jexlExpression = null;

    private static List<String> groupBy(List<Map<String, String>> mapList, String key, List<String> fields) {
        List<String> result = new ArrayList<>();
        List<String> temp = new ArrayList<>();
        for (int i = 0; i < mapList.size(); i++) {
            String s = "";
            for (int j = 0; j < fields.size(); j++) {
                s += mapList.get(i).get(fields.get(j)) + ",";
            }
            boolean b = true;
            for (int k = 0; k < temp.size(); k++) {
                if (temp.get(k).equals(s)) b = false;
            }
            if (b) {
                temp.add(s);
                result.add(mapList.get(i).get(key));
            }
        }
        return result;
    }

    private static List<String> groupBy(List<Map<String, String>> mapList, String key) {
        Set<String> set = new LinkedHashSet<>();
        List<String> list = new ArrayList<>();
        for (int i = 0; i < mapList.size(); i++) {
            set.add(mapList.get(i).get(key));
        }
        Iterator<String> iterator = set.iterator();
        while (iterator.hasNext()) {
            list.add(iterator.next());
        }
        return list;
    }

    private static boolean isCanCal(String s) {
        boolean b = false;
        String[] str = {"+", "-", "*", "/", "%"};
        for (int i = 0; i < str.length; i++) {
            if (s.contains(str[i])) b = true;
        }
        return b;
    }

    private static Double cal(Map<String, String> map, String field) {
        if (field == null) return 0.0;
        for (Map.Entry<String, String> entry : map.entrySet()) {
            if (field.contains(entry.getKey())) {
                field = field.replaceAll(entry.getKey(), map.get(entry.getKey()));
            }
        }
        jexlExpression = jexlEngine.createExpression(field);
        String evaluate = jexlExpression.evaluate(null).toString();
        Double d = Double.parseDouble(evaluate);
        return d;
    }

    public static List<String> sum(List<Map<String, String>> mapList, String field) {
        List<String> list = new ArrayList<>();
        Double sum = 0.0;
        for (int i = 0; i < mapList.size(); i++) {
            Double d = cal(mapList.get(i), field);
            sum += d;
        }
        list.add(String.valueOf(sum));
        return list;
    }

    public static List<String> avg(List<Map<String, String>> mapList, String field) {
        List<String> list = new ArrayList<>();
        Double sum = Double.parseDouble(sum(mapList, field).get(0));
        if (mapList.size() != 0) {
            list.add(String.valueOf(sum / mapList.size()));
        } else {
            list.add("0");
        }
        return list;
    }

    private static int choose(String[] fil, int i, int j, List<Map<String, String>> mapList, int sp) throws ExpException {
        if (fil[j].contains("==")) {
            if ((mapList.get(i).get(fil[j].split("==")[0]) == null || fil[j].split("==")[1] == null))
                throw new ExpException("parsering error in : '" + fil[j] + "'");
            if (mapList.get(i).get(fil[j].split("==")[0]).equals(fil[j].split("==")[1])) {
                sp++;
            }
        } else if (fil[j].contains(">") && !fil[j].contains(">=")) {
            if ((mapList.get(i).get(fil[j].split(">")[0]) == null || fil[j].split(">")[1] == null))
                throw new ExpException("parsering error in : '" + fil[j] + "'");
            if (Double.parseDouble(mapList.get(i).get(fil[j].split(">")[0])) > Double.parseDouble(fil[j].split(">")[1])) {
                sp++;
            }
        } else if (fil[j].contains("<") && !fil[j].contains("<=")) {
            if ((mapList.get(i).get(fil[j].split("<")[0]) == null || fil[j].split("<")[1] == null))
                throw new ExpException("parsering error in : '" + fil[j] + "'");
            if (Double.parseDouble(mapList.get(i).get(fil[j].split("<")[0])) < Double.parseDouble(fil[j].split("<")[1])) {
                sp++;
            }
        } else if (fil[j].contains(">=")) {
            if ((mapList.get(i).get(fil[j].split(">=")[0]) == null || fil[j].split(">=")[1] == null))
                throw new ExpException("parsering error in : '" + fil[j] + "'");
            if (Double.parseDouble(mapList.get(i).get(fil[j].split(">=")[0])) >= Double.parseDouble(fil[j].split(">=")[1])) {
                sp++;
            }
        } else if (fil[j].contains("<=")) {
            if ((mapList.get(i).get(fil[j].split("<=")[0]) == null || fil[j].split("<=")[1] == null))
                throw new ExpException("parsering error in : '" + fil[j] + "'");
            if (Double.parseDouble(mapList.get(i).get(fil[j].split("<=")[0])) <= Double.parseDouble(fil[j].split("<=")[1])) {
                sp++;
            }
        }
        return sp;
    }

    private static List<String> baseAnalysis(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        List<String> list = new ArrayList<>();
        String[] fils = filter.split("&");
        List<String> fill = new ArrayList<>();
        for (int i = 0; i < fils.length; i++) {
            if (!fils[i].equals("")) {
                fill.add(fils[i]);
            }
        }
        String[] fil = new String[fill.size()];
        for (int i = 0; i < fill.size(); i++) {
            fil[i] = fill.get(i);
        }
        for (int i = 0; i < mapList.size(); i++) {
            int sp = 0;
            for (int j = 0; j < fil.length; j++) {
                sp = choose(fil, i, j, mapList, sp);
                if (sp == fil.length) {
                    if (isCanCal(field)) {
                        list.add(String.valueOf(cal(mapList.get(i), field)));
                    } else {
                        list.add(mapList.get(i).get(field));
                    }
                }
            }
        }
        return list;
    }

    public static List<String> sum(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        List<String> list = baseAnalysis(mapList, field, filter);
        double sum = 0.0;
        for (int i = 0; i < list.size(); i++) {
            sum += Double.parseDouble(list.get(i));
        }
        list = new ArrayList<>();
        list.add(String.valueOf(sum));
        return list;
    }

    public static List<String> avg(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        List<String> list = baseAnalysis(mapList, field, filter);
        List<String> avg = new ArrayList<>();
        double sum = 0.0;
        for (int i = 0; i < list.size(); i++) {
            sum += Double.parseDouble(list.get(i));
        }
        if (list.size() == 0) {
            avg.add("0");
        } else {
            avg.add(String.valueOf(sum / list.size()));
        }
        return avg;
    }

    public static List<String> count(List<Map<String, String>> mapList, String filter) throws ExpException {
        List<String> list = baseAnalysis(mapList, null, filter);
        List<String> count = new ArrayList<>();
        count.add(String.valueOf(list.size()));
        return count;
    }

    private static void doCal(String field, List<Map<String, String>> mapList, List<String> list) {
        if (isCanCal(field)) {
            for (int i = 0; i < mapList.size(); i++) {
                Double d = cal(mapList.get(i), field);
                list.add(String.valueOf(d));
            }
        } else {
            for (int i = 0; i < mapList.size(); i++) {
                list.add(mapList.get(i).get(field));
            }
        }
    }

    public static List<String> select(List<Map<String, String>> mapList, String field) {
        List<String> list = new ArrayList<>();
        if (field.contains(":")) {
            if (field.split(":").length == 2 && field.split(":")[1].equals("-1")) {
                field = field.split(":")[0];
                doCal(field, mapList, list);
                Collections.reverse(list);
                return list;
            } else {
                if (isCanCal(field)) {
                    for (int i = 0; i < mapList.size(); i++) {
                        Double d = cal(mapList.get(i), field.split(":")[0]);
                        list.add(String.valueOf(d));
                    }
                } else {
                    for (int i = 0; i < mapList.size(); i++) {
                        list.add(mapList.get(i).get(field.split(":")[0]));
                    }
                }

                return list;
            }
        } else {
            doCal(field, mapList, list);
            return list;
        }
    }

    public static List<String> select(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        if (field.contains(":")) {
            if (field.split(":").length == 2 && field.split(":")[1].equals("-1")) {
                field = field.split(":")[0];
                List<String> list = baseAnalysis(mapList, field, filter);
                Collections.reverse(list);
                return list;
            } else {
                field = field.split(":")[0];
                List<String> list = baseAnalysis(mapList, field, filter);
                return list;
            }
        } else {
            List<String> list = baseAnalysis(mapList, field, filter);
            return list;
        }
    }

    private static List<String> chooseSort(String field, List<Map<String, String>> mapList) {
        List<String> list;
        if (field.split(":").length == 2 && field.split(":")[1].equals("-1")) {
            list = groupBy(mapList, field.split(":")[0]);
            Collections.reverse(list);
            return list;
        } else return groupBy(mapList, field.split(":")[0]);
    }

    public static List<String> group(List<Map<String, String>> mapList, String field) {
        if (field.contains(":")) {
            return chooseSort(field, mapList);
        } else {
            return groupBy(mapList, field);
        }
    }

    public static List<String> group(List<Map<String, String>> mapList, String key, List<String> fields) {
        if (key.contains(":")) {
            return chooseSort(key, mapList);
        } else {
            return groupBy(mapList, key, fields);
        }
    }

    private static List<Map<String, String>> grouping(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        List<Map<String, String>> newMapList = new ArrayList<>();
        String[] fils = filter.split("&");
        List<String> fill = new ArrayList<>();
        for (int i = 0; i < fils.length; i++) {
            if (!fils[i].equals("")) {
                fill.add(fils[i]);
            }
        }
        String[] fil = new String[fill.size()];
        for (int i = 0; i < fill.size(); i++) {
            fil[i] = fill.get(i);
        }
        for (int i = 0; i < mapList.size(); i++) {
            int sp = 0;
            for (int j = 0; j < fil.length; j++) {
                sp = choose(fil, i, j, mapList, sp);
                if (sp == fil.length) {
                    newMapList.add(mapList.get(i));
                }
            }
        }
        return newMapList;
    }

    public static List<String> group(List<Map<String, String>> mapList, String field, String filter) throws ExpException {
        List<Map<String, String>> newMapList = grouping(mapList, field, filter);
        return group(newMapList, field);
    }

    public static List<String> group(List<Map<String, String>> mapList, String key, List<String> fields, String filter) throws ExpException {
        List<Map<String, String>> newMapList = grouping(mapList, key, filter);
        return group(newMapList, key, fields);
    }
}
