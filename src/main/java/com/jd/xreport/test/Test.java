package com.jd.xreport.test;

import com.jd.xreport.analysis.ExcelDataParser;
import com.jd.xreport.exception.DSException;
import com.jd.xreport.exception.ExpException;
import com.jd.xreport.models.DSContainer;
import com.jd.xreport.models.DataSet;
import com.jd.xreport.models.ExcelData;
import com.jd.xreport.utils.ExcelUtil;

import java.io.IOException;
import java.util.*;

/**
 * Created by yuanguangxin on 2017/8/16.
 */
public class Test {
    public static void main(String[] args) throws DSException, IOException, ExpException {
        String path = "C:\\Users\\yuanguangxin\\Desktop\\XReport\\src\\main\\resources\\text.xlsx";
        DataSet dataSet = new DataSet("ds1", "select * from t3");
        DSContainer.addDataSet(dataSet);
        ExcelDataParser ep = new ExcelDataParser(ExcelUtil.readExcel(path));
        ep.parser();
        List<List<ExcelData>> parsedData = ep.getParsedData();
        ExcelUtil.writeExcel("d:\\c.xlsx", parsedData);
    }
}
