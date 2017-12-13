package test;

import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import power.PowerVo;

public class Test {  
    public static void createANewFileTest() {  
        WordBean wordBean = new WordBean();  
        // word.openWord(true);// �� word ����  
        wordBean.setVisible(true);  
        wordBean.createNewDocument();// ����һ�����ĵ�  
        wordBean.setLocation();// ���ô򿪺󴰿ڵ�λ��  
        wordBean.insertText("���");// ���ĵ��в����ַ�  
        wordBean.insertJpeg("D:" + File.separator + "a.jpg"); // ����ͼƬ  
        // ��� ���뱣���ļ�����������  
        // word.saveFileAs("d://a.doc");  
        // word.closeDocument();  
        // word.closeWord();  
    }  
    public static void openAnExistsFileTest() {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // �Ƿ�ǰ̨��word ���򣬻��ߺ�̨����  
        wordBean.openFile("d://a.doc");  
        wordBean.insertJpeg("D:" + File.separator + "a.jpg"); // ����ͼƬ(ע��մ򿪵�word  
        // ����괦�ڿ�ͷ���ʣ�ͼƬ����ǰ������)  
        wordBean.save();  
        wordBean.closeDocument();  
        wordBean.closeWord();  
    }  
    public static void insertFormatStr(String str) {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // �Ƿ�ǰ̨��word ���򣬻��ߺ�̨����  
        wordBean.createNewDocument();// ����һ�����ĵ�  
        wordBean.insertFormatStr(str);// ����һ�����䣬�����е��������������  
    }  
    public static void insertTableTest() {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // �Ƿ�ǰ̨��word ���򣬻��ߺ�̨����  
        wordBean.createNewDocument();// ����һ�����ĵ�  
        wordBean.setLocation();  
        wordBean.insertTable("����", 3, 2);  
        wordBean.saveFileAs("d://table.doc");  
        wordBean.closeDocument();  
        wordBean.closeWord();  
    }  
    public static void mergeTableCellTest() {  
        insertTableTest();//����d://table.doc  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // �Ƿ�ǰ̨��word ���򣬻��ߺ�̨����  
        wordBean.openFile("d://1.doc");  
        wordBean.mergeCellTest();  
    }  
    public static void main(String[] args) throws ClassNotFoundException, SQLException {  
        // ���в���ǰҪ��֤d://a.jpg ͼƬ�ļ�����  
        // createANewFileTest();//����һ�����ļ�  
        // openAnExistsFileTest();// ��һ������ ���ļ�  
        // insertFormatStr("��ʽ ���ַ���");//���ַ�������һ��������  
        //insertTableTest();// ����һ�����  
      // mergeTableCellTest();// �Ա���еĵ�Ԫ����кϲ�  
/*	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=E://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
	    String sql = ("select ��ѵ�Ա��,��ҵ���,��Ŀ���,��Ŀ����,��ѵʱ��,��Ŀ��������,��Ŀ����������,��ѵ����,����  from ������ѵ�ƻ���");  
	    ResultSet rs = stmt.executeQuery(sql);  
	    //stmt.execute(insertSql);
	    while (rs.next()) {  
	        //'System.out.println(rs.getString(1));  
	    	PowerVo Vo = new PowerVo();
	        Vo.setTrainingID(rs.getString(1));  
			Vo.setProjectID(rs.getString(2));
			Vo.setCompanyID(rs.getString(3));
			Vo.setProjectName(rs.getString(4));
			Vo.setTrainingTime(rs.getString(5));
			Vo.setDept(rs.getString(6));
			Vo.setProjectWorkDutyPeople(rs.getString(7));
			Vo.setPeriod(rs.getString(9));
			//Vo.setPeriod(rs.getString(9)+"/"+rs.getString(8));
	       // resultList.add(Vo);
	    }  */
    }  
}
