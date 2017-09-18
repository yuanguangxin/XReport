package com.jd.xreport.models;

import lombok.Getter;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by yuanguangxin on 2017/8/21.
 * <p>
 * 单元格数据类
 */
@Getter
public class ExcelData {
    private String value;//单元格的值
    private int colSpan = 1;//单元格跨几列
    private int rowSpan = 1;//单元格跨几行
    private String exp;//单元格限制条件表达式
    private ExcelData mergeData;//单元格所属合并格
    private ExcelData parent;//单元格父节点
    private ExcelData exParent;//单元格另一个父节点
    private List<ExcelData> children = new ArrayList<>();//单元格孩子集
    private List<ExcelData> exChildren = new ArrayList<>();//单元格另一个孩子集
    private List<List<ExcelData>> parserData = new ArrayList<>();//单元格解析后的单元格结果集

    public ExcelData() {
    }

    public ExcelData(ExcelData ed) {
        this.value = ed.getValue();
        this.colSpan = ed.getColSpan();
        this.rowSpan = ed.getRowSpan();
    }

    public void addChild(ExcelData excelData) {
        this.children.add(excelData);
    }

    public void addExChild(ExcelData excelData) {
        this.exChildren.add(excelData);
    }

    public static ExcelData getExpParent(ExcelData data, String dataSet) {
        ExcelData parent = data.getParent();
        if (dataSet == null) {
            if (parent == null || (parent.getValue().contains("group") && parent.getValue().contains("$"))) {
                return parent;
            } else {
                parent = getExpParent(parent, dataSet);
            }
            return parent;
        } else {
            if (parent == null || (parent.getValue().contains("group") && parent.getValue().contains("$") && parent.getDataSet() != null && parent.getDataSet().equals(dataSet))) {
                return parent;
            } else {
                parent = getExpParent(parent, dataSet);
            }
            return parent;
        }
    }

    public static ExcelData getExpExParent(ExcelData data, String dataSet) {
        ExcelData parent = data.getExParent();
        if (dataSet == null) {
            if (parent == null || (parent.getValue().contains("group") && parent.getValue().contains("#"))) {
                return parent;
            } else {
                parent = getExpExParent(parent, dataSet);
            }
            return parent;
        } else {
            if (parent == null || (parent.getValue().contains("group") && parent.getValue().contains("#") && parent.getDataSet() != null && parent.getDataSet().equals(dataSet))) {
                return parent;
            } else {
                parent = getExpExParent(parent, dataSet);
            }
            return parent;
        }
    }

    public String getDataSet() {
        if (isText()) return null;
        String ds = this.value.substring(1, this.value.indexOf("."));
        return ds;
    }

    public String getField() {
        if (!this.value.contains("group")) {
            return null;
        }
        int index = this.value.indexOf("(");
        String filed = this.value.substring(index + 1, this.value.length() - 1).split(",")[0].split(":")[0];
        return filed;
    }

    public String getFilter() {
        if (isText()) return "";
        int index = this.value.indexOf("(");
        String[] filed = this.value.substring(index + 1, this.value.length() - 1).split(",");
        if (filed.length == 1) return "";
        else return filed[1];
    }

    public void modValue(String s) {
        if (s.equals("") || this.value == null || (!this.value.contains("$") && !this.value.contains("#") && !this.value.contains("~")))
            return;
        int index = this.value.indexOf("(");
        String[] filed = this.value.substring(index + 1, this.value.length() - 1).split(",");
        if (filed.length == 1) {
            this.value = this.value.substring(0, this.value.length() - 1) + "," + s + ")";
        } else {
            this.value = this.value.split(",")[0] + "," + s + ")";
        }
    }


    public void setValue(String value) {
        if (value != null && (value.contains("$") || value.contains("#"))) {
            if (!value.contains("(")) {
                String s = value.substring(0, 1);
                String left = value.substring(1, value.length()).split("\\.")[0];
                String right = value.substring(1, value.length()).split("\\.")[1];
                value = s + left + ".select(" + right + ")";
            }
        }
        this.value = value;
    }

    public boolean isText() {
        if (this.value == null) return true;
        if (!this.value.contains("$") && !this.value.contains("#") && !this.value.contains("~")) {
            return true;
        } else return false;
    }

    public static int getMaxRow(ExcelData ed) {
        if (ed.getChildren().size() == 0) return 1;
        int num = 0;
        for (int i = 0; i < ed.getChildren().size(); i++) {
            num += getMaxRow(ed.getChildren().get(i));
        }
        return num;
    }

    public static int getMaxCol(ExcelData ed) {
        if (ed.getExChildren().size() == 0) return 1;
        int num = 0;
        for (int i = 0; i < ed.getExChildren().size(); i++) {
            num += getMaxCol(ed.getExChildren().get(i));
        }
        return num;
    }

    public void setColSpan(int colSpan) {
        this.colSpan = colSpan;
    }

    public void setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
    }

    public void setExp(String exp) {
        this.exp = exp;
    }

    public void setMergeData(ExcelData mergeData) {
        this.mergeData = mergeData;
    }

    public void setParent(ExcelData parent) {
        this.parent = parent;
    }

    public void setChildren(List<ExcelData> children) {
        this.children = children;
    }

    public void setExChildren(List<ExcelData> exChildren) {
        this.exChildren = exChildren;
    }

    public void setParserData(List<List<ExcelData>> parserData) {
        this.parserData = parserData;
    }

    public void setExParent(ExcelData exParent) {
        this.exParent = exParent;
    }

    @Override
    public String toString() {
        return "ExcelData{" +
                "value='" + value + '\'' +
                ", colSpan=" + colSpan +
                ", rowSpan=" + rowSpan +
                ", exp=" + exp +
                '}';
    }
}
