package com.jd.xreport.analysis;

import com.jd.xreport.exception.DSException;
import com.jd.xreport.exception.ExpException;
import com.jd.xreport.models.DSExpression;
import com.jd.xreport.models.ExcelData;
import com.jd.xreport.models.Point;
import com.jd.xreport.utils.ExcelDataUtil;
import com.jd.xreport.utils.ExcelUtil;
import com.jd.xreport.utils.ListUtil;
import lombok.Data;

import java.io.IOException;
import java.util.*;

/**
 * Created by yuanguangxin on 2017/8/18.
 * <p>
 * 数据解析类
 */
@Data
public class ExcelDataParser {
    private List<List<ExcelData>> data;
    private List<List<ExcelData>> parsedData = new ArrayList<>();

    public ExcelDataParser(List<List<ExcelData>> data) {
        this.data = data;
    }

    private void parserParent() {
        for (int i = 0; i < data.size(); i++) {
            for (int j = 0; j < data.get(i).size(); j++) {
                if (data.get(i).get(j).getValue() == null) {
                    continue;
                }
                if (j == 0) {
                    data.get(i).get(j).setParent(null);
                    continue;
                }
                ExcelData ed = data.get(i).get(j - 1);
                if (ed.getValue() == null) {
                    ed = data.get(i).get(j - 1).getMergeData();
                }
                ed.addChild(data.get(i).get(j));
                data.get(i).get(j).setParent(ed);
            }
        }
    }

    private void parserExParent() {
        for (int i = 0; i < data.size(); i++) {
            for (int j = 0; j < data.get(i).size(); j++) {
                if (data.get(i).get(j).getValue() == null) {
                    continue;
                }
                if (i == 0) {
                    data.get(i).get(j).setExParent(null);
                    continue;
                }
                ExcelData ed = data.get(i - 1).get(j);
                if (ed.getValue() == null) {
                    ed = data.get(i - 1).get(j).getMergeData();
                }
                ed.addExChild(data.get(i).get(j));
                data.get(i).get(j).setExParent(ed);
            }
        }
    }

    private void fillText(List<String> result, ExcelData ed, List<List<ExcelData>> lists, String f) {
        for (int k = 0; k < result.size(); k++) {
            List<ExcelData> dt = new ArrayList<>();
            ExcelData ned = new ExcelData();
            String ex = "";
            if (ed.getField() != null) {
                ex = ed.getField() + "==" + result.get(k);
            }
            if (!f.equals("")) ex = ex + "&" + f;
            ned.setValue(result.get(k));
            ned.setExp(ex);
            dt.add(ned);
            lists.add(dt);
        }
    }

    private void fillLed(List<List<ExcelData>> led, ExcelData ed) {
        List<List<ExcelData>> lists = new ArrayList<>();
        List<ExcelData> eds = new ArrayList<>();
        for (int m = 0; m < led.get(0).size(); m++) {
            ExcelData ned = new ExcelData(ed);
            led.get(0).get(m).addExChild(ned);
            ned.setExParent(led.get(0).get(m));
            eds.add(ned);
        }
        lists.add(eds);
        ed.setParserData(lists);
    }

    private void fillExp(List<List<ExcelData>> led, List<List<ExcelData>> expLed, ExcelData ed, List<ExcelData> edta, String temp) throws ExpException, DSException {
        DSExpression exp;
        for (int m = 0; m < led.get(0).size(); m++) {
            String s = "";
            String pf = "";
            if (expLed != null && expLed.get(0).get(m).getExp() != null) {
                s = expLed.get(0).get(m).getExp();
                pf = expLed.get(0).get(m).getFilter();
            }
            String filter = ed.getFilter();
            if (!pf.equals("")) {
                filter = pf + "&" + filter;
            }
            String f;
            if (filter.equals("")) {
                f = s;
            } else {
                f = filter + "&" + s;
            }
            ed.modValue(f);
            exp = new DSExpression(ed.getValue());
            exp.analysis();
            List<String> result = exp.getResult();
            for (int k = 0; k < result.size(); k++) {
                String nex = f;
                if (ed.getField() != null) {
                    nex = nex + "&" + ed.getField() + "==" + result.get(k);
                }
                ExcelData ned = new ExcelData();
                ned.setValue(result.get(k));
                ned.setExParent(led.get(0).get(m));
                ned.setColSpan(ed.getColSpan());
                led.get(0).get(m).addExChild(ned);
                ned.setExp(nex);
                edta.add(ned);
            }
            ed.setValue(temp);
        }
    }

    private void fillResult(List<String> result, List<List<ExcelData>> led, ExcelData ed, List<List<ExcelData>> lists, String f, int m) {
        for (int k = 0; k < result.size(); k++) {
            List<ExcelData> edta = new ArrayList<>();
            String nex = f;
            if (ed.getField() != null) {
                nex = nex + "&" + ed.getField() + "==" + result.get(k);
            }
            ExcelData ned = new ExcelData();
            ned.setValue(result.get(k));
            ned.setParent(led.get(m).get(0));
            ned.setColSpan(ed.getColSpan());
            led.get(m).get(0).addChild(ned);
            ned.setExp(nex);
            edta.add(ned);
            lists.add(edta);
        }
    }

    private void fillEd(List<List<ExcelData>> led, ExcelData ed, List<List<ExcelData>> lists) {
        for (int m = 0; m < led.size(); m++) {
            List<ExcelData> eds = new ArrayList<>();
            ExcelData ned = new ExcelData(ed);
            led.get(m).get(0).addChild(ned);
            ned.setParent(led.get(m).get(0));
            eds.add(ned);
            lists.add(eds);
        }
    }

    private void fillLists(List<List<ExcelData>> led, List<List<ExcelData>> expLed, ExcelData ed, List<List<ExcelData>> lists, String temp) throws ExpException, DSException {
        DSExpression exp;
        for (int m = 0; m < led.size(); m++) {
            String s = "";
            String pf = "";
            if (expLed != null && expLed.get(m).get(0).getExp() != null) {
                s = expLed.get(m).get(0).getExp();
                pf = expLed.get(m).get(0).getFilter();
            }
            String filter = ed.getFilter();
            if (!pf.equals("")) {
                filter = pf + "&" + filter;
            }
            String f;
            if (filter.equals("")) {
                f = s;
            } else {
                f = filter + "&" + s;
            }
            ed.modValue(f);
            exp = new DSExpression(ed.getValue());
            exp.analysis();
            List<String> result = exp.getResult();
            fillResult(result, led, ed, lists, f, m);
            ed.setValue(temp);
        }
    }

    private void parserData() throws ExpException, DSException {
        for (int i = 0; i < data.size(); i++) {
            for (int j = 0; j < data.get(i).size(); j++) {
                ExcelData ed = data.get(i).get(j);
                DSExpression exp;
                if (ed.getValue() == null) {
                    continue;
                }
                ExcelData parent = ed.getParent();
                ExcelData exParent = ed.getExParent();
                ExcelData expParent;
                ExcelData expExParent;
                if (ed.isText()) {
                    expParent = ExcelData.getExpParent(ed, null);
                    expExParent = ExcelData.getExpExParent(ed, null);
                } else {
                    expParent = ExcelData.getExpParent(ed, ed.getDataSet());
                    expExParent = ExcelData.getExpExParent(ed, ed.getDataSet());
                }
                if (parent == null) {
                    if (exParent == null) {
                        if (ed.isText()) {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            ExcelData ned = new ExcelData(ed);
                            List<ExcelData> eds = new ArrayList<>();
                            eds.add(ned);
                            lists.add(eds);
                            ed.setParserData(lists);
                            continue;
                        }
                        exp = new DSExpression(ed.getValue());
                        exp.analysis();
                        List<String> result = exp.getResult();
                        List<List<ExcelData>> lists = new ArrayList<>();
                        String f = ed.getFilter();
                        if (ed.getValue().contains("#")) {
                            List<ExcelData> dt = new ArrayList<>();
                            for (int k = 0; k < result.size(); k++) {
                                ExcelData ned = new ExcelData();
                                String ex = "";
                                if (ed.getField() != null) {
                                    ex = ed.getField() + "==" + result.get(k);
                                }
                                if (!f.equals("")) ex = ex + "&" + f;
                                ned.setValue(result.get(k));
                                ned.setExp(ex);
                                dt.add(ned);
                            }
                            lists.add(dt);
                            ed.setParserData(lists);
                        } else {
                            fillText(result, ed, lists, f);
                            ed.setParserData(lists);
                        }
                        continue;
                    }
                    if (expExParent == null) {
                        if (ed.isText()) {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            ExcelData ned = new ExcelData(ed);
                            if (exParent.isText()) {
                                ned.setExParent(exParent.getParserData().get(0).get(0));
                                exParent.getParserData().get(0).get(0).addExChild(ned);
                            }
                            List<ExcelData> eds = new ArrayList<>();
                            eds.add(ned);
                            lists.add(eds);
                            ed.setParserData(lists);
                            continue;
                        }
                        exp = new DSExpression(ed.getValue());
                        exp.analysis();
                        List<String> result = exp.getResult();
                        List<List<ExcelData>> lists = new ArrayList<>();
                        String f = ed.getFilter();
                        if (ed.getValue().contains("#")) {
                            List<ExcelData> dt = new ArrayList<>();
                            for (int k = 0; k < result.size(); k++) {
                                ExcelData ned = new ExcelData();
                                String ex = "";
                                if (ed.getField() != null) {
                                    ex = ed.getField() + "==" + result.get(k);
                                }
                                if (!f.equals("")) ex = ex + "&" + f;
                                ned.setValue(result.get(k));
                                ned.setExp(ex);
                                ned.setExParent(exParent.getParserData().get(0).get(0));
                                exParent.getParserData().get(0).get(0).addExChild(ned);
                                dt.add(ned);
                            }
                            lists.add(dt);
                            ed.setParserData(lists);
                        } else {
                            fillText(result, ed, lists, f);
                            ed.setParserData(lists);
                        }
                        continue;
                    }
                    List<List<ExcelData>> led = exParent.getParserData();
                    List<List<ExcelData>> expLed = expExParent.getParserData();
                    String temp = ed.getValue();
                    if (ed.isText()) {
                        fillLed(led, ed);
                        continue;
                    }
                    List<List<ExcelData>> lists = new ArrayList<>();
                    List<ExcelData> edta = new ArrayList<>();
                    fillExp(led, expLed, ed, edta, temp);
                    lists.add(edta);
                    ed.setParserData(lists);
                } else if (exParent == null) {
                    if (expParent == null) {
                        if (ed.isText()) {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            List<ExcelData> eds = new ArrayList<>();
                            ExcelData ned = new ExcelData(ed);
                            eds.add(ned);
                            lists.add(eds);
                            if (parent.isText()) {
                                ExcelData pa = parent.getParserData().get(0).get(0);
                                pa.addChild(ned);
                                ned.setParent(pa);
                            }
                            ed.setParserData(lists);
                            continue;
                        }
                        List<List<ExcelData>> led = parent.getParserData();
                        if (ed.getValue().contains("#")) {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            List<ExcelData> edta = new ArrayList<>();
                            String temp = ed.getValue();
                            for (int m = 0; m < led.size(); m++) {
                                String filter = ed.getFilter();
                                String f = filter;
                                ed.modValue(f);
                                exp = new DSExpression(ed.getValue());
                                exp.analysis();
                                List<String> result = exp.getResult();
                                for (int k = 0; k < result.size(); k++) {
                                    String nex = f;
                                    if (ed.getField() != null) {
                                        nex = nex + "&" + ed.getField() + "==" + result.get(k);
                                    }
                                    ExcelData ned = new ExcelData();
                                    ned.setValue(result.get(k));
                                    ned.setRowSpan(ed.getRowSpan());
                                    ned.setExp(nex);
                                    edta.add(ned);
                                }
                                lists.add(edta);
                                ed.setValue(temp);
                            }
                            ed.setParserData(lists);
                        } else {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            String temp = ed.getValue();
                            for (int m = 0; m < led.size(); m++) {
                                String filter = ed.getFilter();
                                String f = filter;
                                ed.modValue(f);
                                exp = new DSExpression(ed.getValue());
                                exp.analysis();
                                List<String> result = exp.getResult();
                                fillResult(result, led, ed, lists, f, m);
                                ed.setValue(temp);
                            }
                            ed.setParserData(lists);
                        }
                        continue;
                    }
                    List<List<ExcelData>> led = parent.getParserData();
                    List<List<ExcelData>> expLed = expParent.getParserData();
                    String temp = ed.getValue();
                    if (ed.isText()) {
                        List<List<ExcelData>> lists = new ArrayList<>();
                        fillEd(led, ed, lists);
                        ed.setParserData(lists);
                        continue;
                    }
                    List<List<ExcelData>> lists = new ArrayList<>();
                    fillLists(led, expLed, ed, lists, temp);
                    ed.setParserData(lists);
                } else {
                    List<List<ExcelData>> led = parent.getParserData();
                    List<List<ExcelData>> led_ex = exParent.getParserData();
                    if (expParent == null) {
                        if (expExParent == null) {
                            if (ed.isText()) {
                                List<List<ExcelData>> lists = new ArrayList<>();
                                List<ExcelData> eds = new ArrayList<>();
                                ExcelData ned = new ExcelData(ed);
                                if (exParent.isText()) {
                                    led_ex.get(0).get(0).addExChild(ned);
                                    led.get(0).get(0).addChild(ned);
                                }
                                if (parent.isText()) {
                                    ned.setParent(led.get(0).get(0));
                                    ned.setExParent(led_ex.get(0).get(0));
                                }
                                eds.add(ned);
                                lists.add(eds);
                                ed.setParserData(lists);
                                continue;
                            }
                            exp = new DSExpression(ed.getValue());
                            exp.analysis();
                            List<String> result = exp.getResult();
                            List<List<ExcelData>> lists = new ArrayList<>();
                            String f = ed.getFilter();
                            if (ed.getValue().contains("#")) {
                                List<ExcelData> dt = new ArrayList<>();
                                for (int k = 0; k < result.size(); k++) {
                                    ExcelData ned = new ExcelData();
                                    String ex = "";
                                    if (ed.getField() != null) {
                                        ex = ed.getField() + "==" + result.get(k);
                                    }
                                    if (!f.equals("")) ex = ex + "&" + f;
                                    ned.setValue(result.get(k));
                                    ned.setExp(ex);
                                    ned.setExParent(exParent.getParserData().get(0).get(0));
                                    exParent.getParserData().get(0).get(0).addChild(ned);
                                    dt.add(ned);
                                }
                                lists.add(dt);
                                ed.setParserData(lists);
                            } else {
                                for (int k = 0; k < result.size(); k++) {
                                    List<ExcelData> dt = new ArrayList<>();
                                    ExcelData ned = new ExcelData();
                                    String ex = "";
                                    if (ed.getField() != null) {
                                        ex = ed.getField() + "==" + result.get(k);
                                    }
                                    if (!f.equals("")) ex = ex + "&" + f;
                                    ned.setValue(result.get(k));
                                    ned.setExp(ex);
                                    ned.setParent(parent.getParserData().get(0).get(0));
                                    parent.getParserData().get(0).get(0).addChild(ned);
                                    dt.add(ned);
                                    lists.add(dt);
                                }
                                ed.setParserData(lists);
                            }
                            continue;
                        }

                        List<List<ExcelData>> expLed = expExParent.getParserData();
                        String temp = ed.getValue();
                        if (ed.isText()) {
                            fillLed(led_ex, ed);
                            continue;
                        }
                        List<List<ExcelData>> lists = new ArrayList<>();
                        List<ExcelData> edta = new ArrayList<>();
                        fillExp(led_ex, expLed, ed, edta, temp);
                        lists.add(edta);
                        ed.setParserData(lists);
                        continue;
                    }
                    if (expExParent == null) {
                        List<List<ExcelData>> expLed = expParent.getParserData();
                        String temp = ed.getValue();
                        if (ed.isText()) {
                            List<List<ExcelData>> lists = new ArrayList<>();
                            fillEd(led, ed, lists);
                            ed.setParserData(lists);
                            continue;
                        }

                        List<List<ExcelData>> lists = new ArrayList<>();
                        fillLists(led, expLed, ed, lists, temp);
                        ed.setParserData(lists);
                        continue;
                    }
                    List<List<ExcelData>> expLed = expParent.getParserData();
                    List<List<ExcelData>> expLed_ex = expExParent.getParserData();
                    String temp = ed.getValue();
                    if (ed.isText()) {
                        List<List<ExcelData>> lists = new ArrayList<>();
                        for (int m = 0; m < led.size(); m++) {
                            List<ExcelData> eds = new ArrayList<>();
                            for (int n = 0; n < led_ex.get(0).size(); n++) {
                                ExcelData ned = new ExcelData(ed);
                                led_ex.get(0).get(n).addExChild(ned);
                                led.get(m).get(0).addChild(ned);
                                ned.setExParent(led_ex.get(0).get(n));
                                ned.setParent(led.get(m).get(0));
                                eds.add(ned);
                            }
                            lists.add(eds);
                        }
                        ed.setParserData(lists);
                        continue;
                    }
                    List<List<ExcelData>> lists = new ArrayList<>();
                    for (int m = 0; m < led.size(); m++) {
                        List<ExcelData> edta = new ArrayList<>();
                        for (int n = 0; n < led_ex.get(0).size(); n++) {
                            String s = "";
                            String s_ex = "";
                            String pf = "";
                            String pf_ex = "";
                            if (expLed != null && expLed.get(m).get(0).getExp() != null) {
                                s = expLed.get(m).get(0).getExp();
                                pf = expLed.get(m).get(0).getFilter();
                            }
                            if (expLed_ex != null && expLed_ex.get(0).get(n).getExp() != null) {
                                s_ex = expLed_ex.get(0).get(n).getExp();
                                pf_ex = expLed_ex.get(0).get(n).getFilter();
                            }
                            String filter = ed.getFilter();
                            if (!pf.equals("")) {
                                filter = pf + "&" + filter;
                            }
                            if (!pf_ex.equals("")) {
                                filter = filter + "&" + pf_ex;
                            }
                            String f;
                            if (filter.equals("")) {
                                f = s + "&" + s_ex;
                            } else {
                                f = filter + "&" + s + "&" + s_ex;
                            }
                            ed.modValue(f);
                            exp = new DSExpression(ed.getValue());
                            exp.analysis();
                            List<String> result = exp.getResult();
                            for (int k = 0; k < result.size(); k++) {
                                String nex = f;
                                if (ed.getField() != null) {
                                    nex = nex + "&" + ed.getField() + "==" + result.get(k);
                                }
                                ExcelData ned = new ExcelData();
                                ned.setValue(result.get(k));
                                ned.setParent(led.get(m).get(0));
                                ned.setExParent(led_ex.get(0).get(n));
                                ned.setColSpan(ed.getColSpan());
                                led.get(m).get(0).addChild(ned);
                                led_ex.get(0).get(n).addExChild(ned);
                                ned.setExp(nex);
                                edta.add(ned);
                            }
                            ed.setValue(temp);
                        }
                        lists.add(edta);
                    }
                    ed.setParserData(lists);
                }
            }
        }
    }

    public void parser() throws ExpException, DSException, IOException {
        parserParent();
        parserExParent();
        parserData();
        Point ph = ExcelDataUtil.getHorPoint(data);
        Point pv = ExcelDataUtil.getVerPoint(data);
        if (ph.getX() == -1) {
            int lineNum = 0;
            for (int i = 0; i < data.size(); i++) {
                ExcelData ed = data.get(i).get(0);
                if (ed.getValue() == null) continue;
                List<List<ExcelData>> parserData = ed.getParserData();
                for (int k = 0; k < parserData.size(); k++) {
                    int tnum = 0;
                    ExcelData ped = parserData.get(k).get(0);
                    int rowspan = ExcelData.getMaxRow(ped);
                    int colspan = ped.getColSpan();
                    List<ExcelData> list = new ArrayList<>();
                    ped.setRowSpan(rowspan);
                    list.add(ped);
                    for (int a = 1; a < colspan; a++) {
                        ExcelData ned = new ExcelData();
                        ned.setValue(null);
                        list.add(ned);
                        ned.setMergeData(ped);
                    }
                    parsedData.add(list);
                    tnum++;
                    for (int b = 1; b < rowspan; b++) {
                        List<ExcelData> nlist = new ArrayList<>();
                        for (int a = 0; a < colspan; a++) {
                            ExcelData ned = new ExcelData();
                            ned.setValue(null);
                            nlist.add(ned);
                            ned.setMergeData(ped);
                        }
                        parsedData.add(nlist);
                        tnum++;
                    }
                    ergodicDataVer(ped, lineNum);
                    lineNum += tnum;
                }
            }
            return;
        }
        if (pv.getX() == -1) {
            List<ExcelData> list = new ArrayList<>();
            parsedData.add(list);
            for (int i = 0; i < data.get(0).size(); i++) {
                ExcelData ed = data.get(0).get(i);
                if (ed.getValue() == null) continue;
                List<List<ExcelData>> parserData = ed.getParserData();
                for (int k = 0; k < parserData.get(0).size(); k++) {
                    int tnum = 0;
                    ExcelData ped = parserData.get(0).get(k);
                    int rowspan = ped.getRowSpan();
                    int colspan = ExcelData.getMaxCol(ped);
                    ped.setRowSpan(rowspan);
                    ped.setColSpan(colspan);
                    list.add(ped);
                    for (int a = 1; a < colspan; a++) {
                        ExcelData ned = new ExcelData();
                        ned.setValue(null);
                        ned.setMergeData(ped);
                        list.add(ned);
                    }
                    tnum++;
                    for (int b = 1; b < rowspan; b++) {
                        if (parsedData.size() >= b) {
                            for (int a = 0; a < colspan; a++) {
                                ExcelData ned = new ExcelData();
                                ned.setValue(null);
                                ned.setMergeData(ped);
                                parsedData.get(b).add(ned);
                            }
                        } else {
                            List<ExcelData> nlist = new ArrayList<>();
                            for (int a = 0; a < colspan; a++) {
                                ExcelData ned = new ExcelData();
                                ned.setMergeData(ped);
                                ned.setValue(null);
                                nlist.add(ned);
                            }
                            parsedData.add(nlist);
                        }
                        tnum++;
                    }
                    ergodicDataHor(ped, 1);
                }
            }
            return;
        }


        int textWidth = ph.getY();
        int textHeight = pv.getX();
        List<List<ExcelData>> textList = new ArrayList<>();
        for (int i = 0; i < textHeight; i++) {
            List<ExcelData> temp = new ArrayList<>();
            for (int j = 0; j < textWidth; j++) {
                temp.add(data.get(i).get(j));
            }
            textList.add(temp);
        }
        ExcelDataParser tep = new ExcelDataParser(textList);
        tep.parser();
        List<List<ExcelData>> textResult = tep.getParsedData();

        List<List<ExcelData>> verList = new ArrayList<>();
        for (int i = pv.getX(); i < data.size(); i++) {
            List<ExcelData> temp = new ArrayList<>();
            for (int j = pv.getY(); j < data.get(i).size(); j++) {
                temp.add(data.get(i).get(j));
            }
            verList.add(temp);
        }

        ExcelDataParser vep = new ExcelDataParser(verList);
        vep.parser();
        List<List<ExcelData>> verResult = vep.getParsedData();

        List<List<ExcelData>> verChildren = new ArrayList<>();
        for (int i = 0; i < verResult.size(); i++) {
            if (verResult.get(i).get(verResult.get(i).size() - 2).getValue() == null) continue;
            verChildren.add(verResult.get(i).get(verResult.get(i).size() - 2).getChildren());
        }

        List<List<ExcelData>> verList_r = new ArrayList<>();
        for (int i = pv.getX(); i < data.size(); i++) {
            List<ExcelData> temp = new ArrayList<>();
            for (int j = pv.getY(); j < ph.getY(); j++) {
                temp.add(data.get(i).get(j));
            }
            verList_r.add(temp);
        }

        ExcelDataParser vepr = new ExcelDataParser(verList_r);
        vepr.parser();
        List<List<ExcelData>> verResult_r = vepr.getParsedData();

        for (int i = 0; i < verResult_r.size(); i++) {
            if (verResult_r.get(i).get(verResult_r.get(i).size() - 1).getValue() == null) continue;
            verResult_r.get(i).get(verResult_r.get(i).size() - 1).setChildren(verChildren.get(i));
        }

        List<List<ExcelData>> horList = new ArrayList<>();
        for (int i = ph.getX(); i < data.size(); i++) {
            List<ExcelData> temp = new ArrayList<>();
            for (int j = ph.getY(); j < data.get(i).size(); j++) {
                temp.add(data.get(i).get(j));
            }
            horList.add(temp);
        }
        ExcelDataParser hep = new ExcelDataParser(horList);
        hep.parser();
        List<List<ExcelData>> horResult = hep.getParsedData();

        List<List<ExcelData>> horChildren = new ArrayList<>();
        for (int i = 0; i < horResult.get(horResult.size() - 2).size(); i++) {
            if (horResult.get(horResult.size() - 2).get(i).getValue() == null) continue;
            horChildren.add(horResult.get(horResult.size() - 2).get(i).getExChildren());
        }


        List<List<ExcelData>> horList_r = new ArrayList<>();
        for (int i = ph.getX(); i < pv.getX(); i++) {
            List<ExcelData> temp = new ArrayList<>();
            for (int j = ph.getY(); j < data.get(i).size(); j++) {
                temp.add(data.get(i).get(j));
            }
            horList_r.add(temp);
        }
        ExcelDataParser hepr = new ExcelDataParser(horList_r);
        hepr.parser();
        List<List<ExcelData>> horResult_r = hepr.getParsedData();

        for (int i = 0; i < horResult_r.get(horResult_r.size() - 1).size(); i++) {
            if (horResult_r.get(horResult_r.size() - 1).get(i).getValue() == null) continue;
            horResult_r.get(horResult_r.size() - 1).get(i).setExChildren(horChildren.get(i));
        }

        ListUtil.joinTop(textResult, verResult_r);
        ListUtil.joinLeft(textResult, horResult_r);


        boolean bool = true;
        int a = -1, b = -1;
        for (int i = 0; i < textResult.size(); i++) {
            for (int j = 0; j < textResult.get(i).size(); j++) {
                if (textResult.get(i).get(j) == null) {
                    if (bool) {
                        a = i - 1;
                        b = j - 1;
                        bool = false;
                    }
                    ExcelData left = textResult.get(i).get(b);
                    if (left.getValue() == null) left = left.getMergeData();
                    ExcelData top = textResult.get(a).get(j);
                    if (top.getValue() == null) top = top.getMergeData();
                    int bl = 0;
                    for (int m = 0; m < left.getChildren().size(); m++) {
                        for (int n = 0; n < top.getExChildren().size(); n++) {
                            if (top.getExChildren().get(n).getValue().equals(left.getChildren().get(m).getValue())) {
                                textResult.get(i).remove(j);
                                textResult.get(i).add(j, left.getChildren().get(m));
                                bl = 1;
                                break;
                            }
                            if (bl == 1) {
                                break;
                            }
                        }
                    }
                }
            }
        }
        parsedData = textResult;
    }

    private int getGen(ExcelData ed) {
        int gen = 0;
        while (ed.getExParent() != null) {
            gen++;
            ed = ed.getExParent();
        }
        return gen;
    }

    private void ergodicDataHor(ExcelData ed, int rowNum) {
        Stack<ExcelData> stack = new Stack<>();
        stack.push(ed);
        while (!stack.empty()) {
            ExcelData temp = stack.pop();
            int gen = getGen(temp);
            insertDataHor(temp, gen);
            for (int i = temp.getExChildren().size() - 1; i >= 0; i--) {
                stack.push(temp.getExChildren().get(i));
            }
        }
    }

    private void insertDataHor(ExcelData ed, int rowNum) {
        List<ExcelData> children = ed.getExChildren();
        int t = rowNum;
        if (children.size() != 0) {
            rowNum++;
        }
        for (int i = 0; i < children.size(); i++) {
            if (children.get(i).getValue() == null) continue;
            ExcelData ced = children.get(i);
            int rowspan = ced.getRowSpan();
            ced.setRowSpan(rowspan);
            int n = parsedData.get(t).size();
            int st = t;
            for (int k = t; k < parsedData.size(); k++) {
                if (parsedData.get(k).size() < n) {
                    st = k;
                    break;
                }
            }
            if (st == t) {
                st = parsedData.size();
                for (int m = 0; m < rowspan; m++) {
                    List<ExcelData> nedList = new ArrayList<>();
                    parsedData.add(nedList);
                }
            }
            int colspan = ExcelData.getMaxCol(ced);
            ced.setColSpan(colspan);
            parsedData.get(st).add(ced);
            for (int k = 1; k < colspan; k++) {
                ExcelData ned = new ExcelData();
                ned.setMergeData(ced);
                ned.setValue(null);
                parsedData.get(st).add(ned);
            }
            for (int j = 1; j < rowspan; j++) {
                if (parsedData.size() >= st + j) {
                    for (int k = 0; k < colspan; k++) {
                        ExcelData ned = new ExcelData();
                        ned.setMergeData(ced);
                        ned.setValue(null);
                        parsedData.get(st + j).add(ned);
                    }
                } else {
                    List<ExcelData> nlist = new ArrayList<>();
                    for (int k = 0; k < colspan; k++) {
                        ExcelData ned = new ExcelData();
                        ned.setValue(null);
                        ned.setMergeData(ced);
                        nlist.add(ned);
                    }
                    parsedData.add(nlist);
                }
                rowNum++;
            }
        }
    }

    private void insertDataVer(ExcelData ed, int rowNum) {
        List<ExcelData> children = ed.getChildren();
        int t = rowNum;
        for (int i = 0; i < children.size(); i++) {
            if (children.get(i).getValue() == null) continue;
            int n = parsedData.get(t).size();
            int st = t;
            for (int k = t; k < parsedData.size(); k++) {
                if (parsedData.get(k).size() < n) {
                    st = k;
                    break;
                }
            }
            ExcelData ced = children.get(i);
            parsedData.get(st).add(ced);
            int colspan = ced.getColSpan();
            for (int k = 1; k < colspan; k++) {
                ExcelData ned = new ExcelData();
                ned.setValue(null);
                ned.setMergeData(ced);
                parsedData.get(st).add(ned);
            }
            rowNum++;
            int rowspan = ExcelData.getMaxRow(ced);
            ced.setRowSpan(rowspan);
            for (int j = 1; j < rowspan; j++) {
                for (int k = 0; k < colspan; k++) {
                    ExcelData ned = new ExcelData();
                    ned.setValue(null);
                    ned.setMergeData(ced);
                    parsedData.get(st + j).add(ned);
                }
                rowNum++;
            }
        }
    }

    private void ergodicDataVer(ExcelData ed, int rowNum) {
        Queue<ExcelData> queue = new LinkedList<>();
        ExcelData current;
        queue.offer(ed);
        while (!queue.isEmpty()) {
            current = queue.poll();
            int colspan = current.getColSpan();
            List<ExcelData> ec = current.getChildren();
            ExcelData temp = current;
            for (int k = 1; k < colspan; k++) {
                List<ExcelData> ch = new ArrayList<>();
                ExcelData ned = new ExcelData();
                ned.setValue(null);
                ch.add(ned);
                temp.setChildren(ch);
                temp = ned;
            }
            temp.setChildren(ec);
            insertDataVer(current, rowNum);
            for (int i = 0; i < current.getChildren().size(); i++) {
                queue.offer(current.getChildren().get(i));
            }
        }
    }
}

