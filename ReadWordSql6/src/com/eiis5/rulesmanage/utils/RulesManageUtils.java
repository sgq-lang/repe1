package com.eiis5.rulesmanage.utils;

import java.sql.Connection;
import java.sql.DriverManager;

public class RulesManageUtils {
	public static Connection getConnection(String domainname) throws Exception{
		Connection conn = null;
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection("jdbc:oracle:thin:@10.68.19.178:1521:jsbz","EIIS_DEV01","123456");	
			return conn;
		} catch (Exception e) {
			
		}
		return conn;
	}
}
