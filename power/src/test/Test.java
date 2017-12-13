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
        // word.openWord(true);// 打开 word 程序  
        wordBean.setVisible(true);  
        wordBean.createNewDocument();// 创建一个新文档  
        wordBean.setLocation();// 设置打开后窗口的位置  
        wordBean.insertText("你好");// 向文档中插入字符  
        wordBean.insertJpeg("D:" + File.separator + "a.jpg"); // 插入图片  
        // 如果 ，想保存文件，下面三句  
        // word.saveFileAs("d://a.doc");  
        // word.closeDocument();  
        // word.closeWord();  
    }  
    public static void openAnExistsFileTest() {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // 是否前台打开word 程序，或者后台运行  
        wordBean.openFile("d://a.doc");  
        wordBean.insertJpeg("D:" + File.separator + "a.jpg"); // 插入图片(注意刚打开的word  
        // ，光标处于开头，故，图片在最前方插入)  
        wordBean.save();  
        wordBean.closeDocument();  
        wordBean.closeWord();  
    }  
    public static void insertFormatStr(String str) {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // 是否前台打开word 程序，或者后台运行  
        wordBean.createNewDocument();// 创建一个新文档  
        wordBean.insertFormatStr(str);// 插入一个段落，对其中的字体进行了设置  
    }  
    public static void insertTableTest() {  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // 是否前台打开word 程序，或者后台运行  
        wordBean.createNewDocument();// 创建一个新文档  
        wordBean.setLocation();  
        wordBean.insertTable("表名", 3, 2);  
        wordBean.saveFileAs("d://table.doc");  
        wordBean.closeDocument();  
        wordBean.closeWord();  
    }  
    public static void mergeTableCellTest() {  
        insertTableTest();//生成d://table.doc  
        WordBean wordBean = new WordBean();  
        wordBean.setVisible(true); // 是否前台打开word 程序，或者后台运行  
        wordBean.openFile("d://1.doc");  
        wordBean.mergeCellTest();  
    }  
    public static void main(String[] args) throws ClassNotFoundException, SQLException {  
        // 进行测试前要保证d://a.jpg 图片文件存在  
        // createANewFileTest();//创建一个新文件  
        // openAnExistsFileTest();// 打开一个存在 的文件  
        // insertFormatStr("格式 化字符串");//对字符串进行一定的修饰  
        //insertTableTest();// 创建一个表格  
      // mergeTableCellTest();// 对表格中的单元格进行合并  
/*	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=E://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
	    String sql = ("select 培训自编号,企业编号,项目编号,项目名称,培训时间,项目开发部门,项目运行责任人,培训期数,期数  from 短期培训计划表");  
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
