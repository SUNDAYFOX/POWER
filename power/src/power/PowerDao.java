package power;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.apache.struts.action.ActionForm;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import util.Constant;





public class PowerDao {
	public List getClassList(ActionForm form,HttpServletRequest request) throws BiffException, IOException, ClassNotFoundException, SQLException  {
		List resultList = new ArrayList();
		PowerForm theform = (PowerForm) form;
	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	   // String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
	    Statement stmt = conn.createStatement();  
	    String sql = ("select ��ѵ������,��ѵʱ��,���β���,��Ŀ����������,����ר��,��������,��ѵ�Ա��,���� from ��ѵ�������");  
	    ResultSet rs = stmt.executeQuery(sql);  
	    //stmt.execute(insertSql);
	    while (rs.next()) {  
	        //'System.out.println(rs.getString(1));  
	    	PowerVo Vo = new PowerVo();
	        Vo.setClassName(rs.getString(1));  
	        Vo.setTrainingTime(rs.getString(2));
			Vo.setDutyDept(rs.getString(3));
			Vo.setProjectDevelopmentDutyPeople(rs.getString(4));
			Vo.setProjectWorkDutyPeople(rs.getString(5));
			Vo.setStore(rs.getString(6));
			Vo.setTrainingID(rs.getString(7));
			Vo.setPeriod(rs.getString(8));
/*			Vo.setProjectName(rs.getString(4));
			Vo.setTrainingTime(rs.getString(5));
			Vo.setDept(rs.getString(6));*/
			//Vo.setProjectWorkDutyPeople(rs.getString(7));
			//Vo.setPeriod(rs.getString(9));
			//Vo.setPeriod(rs.getString(9)+"/"+rs.getString(8));
	        resultList.add(Vo);
	    }  

	    rs.close();  
	    stmt.close();  
	    conn.close();  
	
		
		
		
		
		
		
/*		// ��������������ȡExcel  
		InputStream is;
		try {
			String path=request.getServletContext().getRealPath("\\DB");
			is = new FileInputStream(new File(path+"\\2017�깫˾������ѵ�ƻ�.xls").getAbsolutePath());
			Workbook wbin = Workbook.getWorkbook(is);  
			// Excel��ҳǩ����  
			int sheet_size = wbin.getNumberOfSheets();  
			//for (int index = 0; index < sheet_size; index++) {  
			// ÿ��ҳǩ����һ��Sheet����  
			Sheet sheet = wbin.getSheet(0);  
			// sheet.getRows()���ظ�ҳ������0 
			System.out.println(theform.getProjectWorkDutyPeople()+"----------------------------------");
			for (int i = 2; i < sheet.getRows(); i++) {  
				String projectWorkDutyPeople = sheet.getCell(21, i).getContents(); 			
				if(theform.getProjectWorkDutyPeople()==null||theform.getProjectWorkDutyPeople().trim().equals("")||projectWorkDutyPeople.indexOf(theform.getProjectWorkDutyPeople())>-1){
					PowerVo Vo = new PowerVo();
					Vo.setTrainingID(sheet.getCell(1, i).getContents());
					Vo.setProjectID(sheet.getCell(2, i).getContents());
					Vo.setCompanyID(sheet.getCell(3, i).getContents());
					Vo.setProjectName(sheet.getCell(4, i).getContents());
					Vo.setTrainingTime(sheet.getCell(11, i).getContents());
					Vo.setDept(sheet.getCell(18, i).getContents());
					Vo.setProjectWorkDutyPeople(sheet.getCell(21, i).getContents());
					resultList.add(Vo);
				}

			}  
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			return resultList;
		}*/
		return resultList;
		

		
		
	}	
}
