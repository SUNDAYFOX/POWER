package power;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class AcessTest {
	public static void main(String[] args) throws ClassNotFoundException, SQLException {
	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	    /** 
	     * 直接连接access文件。 
	     */  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
	   // String insertSql="  insert into 短期培训计划表(培训自编号,企业编号,项目名称,项目编号,培训时间,培训期数,期数 ,项目开发责任人) values('jj2017-01-01','9120001600NU','23746859','安全监督管理培训','10','各供电单位及相关直属单位安监部门安全管理人员',1,'胡炼')";
	  //  String insertSql="  insert into 短期培训计划表(培训自编号,企业编号,项目名称,项目编号,培训时间,培训期数,期数 ,项目开发责任人) values('jj2017-05-02','国网竞赛','23746823','9120001600QK','','电力公司从事无人机巡检专业的在岗职工',1,'彭长均')";
	    //insert into 短期培训计划表(培训自编号,企业编号,项目名称,项目编号,培训时间,培训期数,期数 ,项目开发责任人) values('','','调度自动化专业人员技能提升培训','','4','8','1','周虹')
	    
	  //、 insert into 短期培训计划表(培训自编号,企业编号,项目名称,项目编号,培训时间,培训期数,期数 ,项目开发责任人) values(jj2017-01-01,9120001600NU,安全监督管理培训,23746859,10,1,1,胡炼)
	  ResultSet rs = stmt.executeQuery("select count(1) from 短期培训计划表  where 培训自编号='jj2017-01-01' or 培训自编号='jj2017-01-02'");  
	    //stmt.execute(insertSql);
	    while (rs.next()) {  
	        System.out.println(rs.getString(1));  
	    }  
	    rs.close();   
	    stmt.close();  
	    conn.close();  
	}
}
