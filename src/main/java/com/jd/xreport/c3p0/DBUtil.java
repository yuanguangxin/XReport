package com.jd.xreport.c3p0;

/**
 * Created by yuanguangxin on 2017/8/8.
 *
 * 数据库查询工具类
 */

import lombok.extern.log4j.Log4j;

import java.sql.*;
import java.util.*;

@Log4j
public class DBUtil {

    public static List<Map<String, String>> executeQuery(String sql, Object... o) {
        List<Map<String, String>> list = new ArrayList();
        Connection conn = C3P0Util.getConnection();
        PreparedStatement ps = null;
        ResultSet res = null;
        ResultSetMetaData rsm;
        try {
            ps = conn.prepareStatement(sql);
        } catch (Exception e) {
            log.error("prepare statement failed", e);
            C3P0Util.close(conn, ps, res);
            return null;
        }
        if (o != null) {
            for (int i = 0; i < o.length; i++) {
                try {
                    ps.setObject(i + 1, o[i]);
                } catch (SQLException e) {
                    log.error("prepare statement failed", e);
                    C3P0Util.close(conn, ps, res);
                    return null;
                }
            }
        }
        try {
            res = ps.executeQuery();
            rsm = res.getMetaData();
        } catch (SQLException e) {
            log.error("query failed sql is " + sql, e);
            return null;
        }
        try {
            while (res.next()) {
                Map<String, String> map = new LinkedHashMap();
                for (int i = 0; i < rsm.getColumnCount(); i++) {
                    map.put(rsm.getColumnName(i + 1), String.valueOf(res.getObject(i + 1)));
                }
                list.add(map);
            }
        } catch (Exception e) {
            log.error("execute failed", e);
        }
        C3P0Util.close(conn, ps, res);
        return list;
    }

}
