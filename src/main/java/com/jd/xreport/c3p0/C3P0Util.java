package com.jd.xreport.c3p0;

/**
 * Created by yuanguangxin on 2017/8/8.
 *
 * 连接数据库工具类
 */

import com.mchange.v2.c3p0.ComboPooledDataSource;
import lombok.extern.log4j.Log4j;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

@Log4j
public class C3P0Util {

    static ComboPooledDataSource cpds=null;
    static{
        cpds = new ComboPooledDataSource("mysql");
    }

    public static Connection getConnection(){
        try {
            return cpds.getConnection();
        } catch (SQLException e) {
            log.error(e);
            return null;
        }
    }

    public static void close(Connection conn,PreparedStatement pst,ResultSet rs){
        if(rs!=null){
            try {
                rs.close();
            } catch (SQLException e) {
                log.error(e);
            }
        }
        if(pst!=null){
            try {
                pst.close();
            } catch (SQLException e) {
                log.error(e);
            }
        }

        if(conn!=null){
            try {
                conn.close();
            } catch (SQLException e) {
                log.error(e);
            }
        }
    }
}