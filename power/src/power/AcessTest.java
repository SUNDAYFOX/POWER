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
	     * ֱ������access�ļ��� 
	     */  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
	   // String insertSql="  insert into ������ѵ�ƻ���(��ѵ�Ա��,��ҵ���,��Ŀ����,��Ŀ���,��ѵʱ��,��ѵ����,���� ,��Ŀ����������) values('jj2017-01-01','9120001600NU','23746859','��ȫ�ල������ѵ','10','�����絥λ�����ֱ����λ���ಿ�Ű�ȫ������Ա',1,'����')";
	  //  String insertSql="  insert into ������ѵ�ƻ���(��ѵ�Ա��,��ҵ���,��Ŀ����,��Ŀ���,��ѵʱ��,��ѵ����,���� ,��Ŀ����������) values('jj2017-05-02','��������','23746823','9120001600QK','','������˾�������˻�Ѳ��רҵ���ڸ�ְ��',1,'����')";
	    //insert into ������ѵ�ƻ���(��ѵ�Ա��,��ҵ���,��Ŀ����,��Ŀ���,��ѵʱ��,��ѵ����,���� ,��Ŀ����������) values('','','�����Զ���רҵ��Ա����������ѵ','','4','8','1','�ܺ�')
	    
	  //�� insert into ������ѵ�ƻ���(��ѵ�Ա��,��ҵ���,��Ŀ����,��Ŀ���,��ѵʱ��,��ѵ����,���� ,��Ŀ����������) values(jj2017-01-01,9120001600NU,��ȫ�ල������ѵ,23746859,10,1,1,����)
	  ResultSet rs = stmt.executeQuery("select count(1) from ������ѵ�ƻ���  where ��ѵ�Ա��='jj2017-01-01' or ��ѵ�Ա��='jj2017-01-02'");  
	    //stmt.execute(insertSql);
	    while (rs.next()) {  
	        System.out.println(rs.getString(1));  
	    }  
	    rs.close();   
	    stmt.close();  
	    conn.close();  
	}
}
