package power;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;

import javax.servlet.http.HttpSession;

import jxl.Cell;
import jxl.DateCell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import util.Constant;
import util.ReturnInfo;
import util.Util;

public class PowerExcleDao {
	
	//计划表
	public ReturnInfo importtrainingplanExcel(PowerForm theform, HttpSession session) throws ClassNotFoundException, BiffException, IOException, SQLException {
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			wb = Workbook.getWorkbook(new_is);
		
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=E://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
		    	conn.setAutoCommit(false);
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");

/*				String deleteSql = " delete from 短期培训计划表";
				stmt.execute(deleteSql);*/
				for(int i=1;i<rowSize;i++){	
	 				String trainingID = sheet.getCell(1, i).getContents(); 
	 				String companyID = sheet.getCell(2, i).getContents();
	 				String projectID = sheet.getCell(3, i).getContents();
	 				String projectName = sheet.getCell(4, i).getContents();
	 				String trainingObject = sheet.getCell(5, i).getContents();
	 				String trainingNum = sheet.getCell(6, i).getContents();
	 				String trainingDayNum = sheet.getCell(7, i).getContents();
	 				String projectType = sheet.getCell(16, i).getContents(); 
	 				String projectDevelopmentDept = sheet.getCell(17, i).getContents(); 
	 				String projectDevelopmentPeople = sheet.getCell(18, i).getContents(); 
	 				String projectDevelopmentPeoplePhone = sheet.getCell(19, i).getContents(); 
	 				String projectWorkPeople = sheet.getCell(20, i).getContents(); 		
	 				//String projectWorkPeoplePhone = sheet.getCell(21, i).getContents(); 	
	 				String fee = sheet.getCell(22, i).getContents(); 
	 				String insertSql = "  insert into 培训班基础表(培训项目编号,企业编号,项目名称,项目编号,培训对象,项目开发责任部门,项目开发责任人,开发责任人电话,项目分口,储备费用,计划期数,预计培训天数 ,培训时间,培训班名称,已完成班期数,期数,运行专责) values('"+trainingID+"','"+companyID+"','"+projectName+"','"+projectID+"','"
	 									+trainingObject+"','"+projectDevelopmentDept+"','"+projectDevelopmentPeople+"','"+projectDevelopmentPeoplePhone+"','"+projectType+"','"+fee+"','"+trainingNum+"','"+trainingDayNum+"','";
					String months = sheet.getCell(11, i).getContents().replace("月", "");
					String month[]  =months.split("、"); 
/*					if(trainingID.equals("jj2017-01-94")){
						System.out.println(projectName);
					}*/
					int classNum = 1;
					if(months.indexOf("、")>-1){//5、11月
						for (int i1 = 0; i1 < month.length; i1++) {
							if(month[i1].indexOf("(")>-1){
								int num = Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
								for (int i2 = 1; i2 < num+1; i2++) {
									
									String className = projectName+"(第"+Util.ToNumeric(classNum)+"期)";
									String sql = insertSql+month[i1].substring(0,month[i1].indexOf("("))+"','"+className+"','"+classNum+"','"+classNum+"','"+projectWorkPeople+"')";
									classNum++;
									stmt.execute(sql);
								}
							}else{
								String className = projectName+"(第"+Util.ToNumeric(classNum)+"期)";
								String sql = insertSql+month[i1]+"','"+className+"','"+classNum+"','"+classNum+"','"+projectWorkPeople+"')";	
								classNum++;
								stmt.execute(sql);
								 
							}
						}
					}else{//5(2)月
						if(months.indexOf("(")>-1){
							int num = Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
							for (int i3 = 1; i3 < num+1; i3++) {
								String className = projectName+"(第"+Util.ToNumeric(classNum)+"期)";
								String sql = insertSql+months.substring(0,months.indexOf("("))+"','"+className+"','"+classNum+"','"+classNum+"','"+projectWorkPeople+"')";
								classNum++;
								stmt.execute(sql);
							}
						}else{
							String className = projectName+"(第"+Util.ToNumeric(classNum)+"期)";
							String sql = insertSql+months+"','"+className+"','"+classNum+"','"+classNum+"','"+projectWorkPeople+"')";
							classNum++;
							stmt.execute(sql);
						}     
					}
				}		    	
		    } catch (Exception e) {
				e.printStackTrace();
				conn.rollback();
				returnInfo.setFlag(false);
				return returnInfo;
			}
		
		conn.commit();
        stmt.close();  
        conn.close(); 				    
		returnInfo.setFlag(true);
		return returnInfo;
	}	
	//运行表
	public ReturnInfo importTrainingrunExcel(PowerForm theform, HttpSession session) throws ClassNotFoundException, BiffException, IOException, SQLException {
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			wb = Workbook.getWorkbook(new_is);
		
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  	
		    	conn.setAutoCommit(false);
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");

				
				for(int i=2;i<rowSize;i++){	
					String trainingID = sheet.getCell(1, i).getContents(); 
					String calssID = sheet.getCell(2, i).getContents(); 
					String calssName = sheet.getCell(3, i).getContents(); 
					String num = sheet.getCell(4, i).getContents(); 
					String numInRoom = sheet.getCell(5, i).getContents(); 
					String projectDutyPeople = sheet.getCell(6, i).getContents(); 
					String projectWorkDutyPeople = sheet.getCell(7, i).getContents(); 
					String trainingPlace = sheet.getCell(8, i).getContents(); 
					DateCell start = (DateCell)sheet.getCell(9, i);			    
				    String startTime = format.format(start.getDate());
				      // String sql = "INSERT INTO 培训班运行表(开班时间) VALUES (#" + d + "#)";
					//String endTime = sheet.getCell(10, i).getContents(); 
					DateCell end = (DateCell)sheet.getCell(10, i);			    
				    String endTime = format.format(end.getDate());
					String classRoom = sheet.getCell(11, i).getContents(); 
					String isEveningStudy = sheet.getCell(12, i).getContents(); 
					String Headmaster = sheet.getCell(13, i).getContents(); 
					String remark = sheet.getCell(14, i).getContents(); 
					String insertSql = "  insert into 培训班运行表(培训项目编号,培训班编号,培训班名称,人数,住宿人数,项目责任人,运行专责,培训地点,开班时间,结束时间,地点,是否晚自习,班主任,备注,培训时间,期数) values('"
					+trainingID+"','"+calssID+"','"+calssName+"','"+num+"','"+numInRoom+"','"+projectDutyPeople+"','"+projectWorkDutyPeople+"','"+trainingPlace+"',#"+startTime+"#,#"
					+endTime+"#,'"+classRoom+"','"+isEveningStudy+"','"+Headmaster+"','"+remark+"','"+theform.getTrainingTime()+"','"+theform.getPeriod()+"')";
					stmt.execute(insertSql);
				}		    	
		    } catch (Exception e) {
				e.printStackTrace();
				conn.rollback();
				returnInfo.setFlag(false);
				return returnInfo;
			}
		
		conn.commit();
        stmt.close();  
        conn.close(); 				    
		returnInfo.setFlag(true);
		return returnInfo;
	}
	public ReturnInfo importtrainingcourseExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();
				SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
				
				Range[] rangeCell = sheet.getMergedCells();
				for(int i=2;i<rowSize;i++){	
					Cell cell = sheet.getCell(0, i);
					System.out.println(cell.getType());
					//DateCell day = (DateCell)sheet.getCell(0, i);			    
				    //String dayTime = format.format(day.getDate());		
					//String day = sheet.getCell(0, i).getContents(); 
					int k = i;
					while(cell.getType().toString().equals("Empty")){
						cell = sheet.getCell(0,k--);
						//day = sheet.getCell(0, k--).getContents(); 
					}
					String dayTime = format.format(((DateCell)cell).getDate());
					int j = i;
					String noon = sheet.getCell(1, i).getContents(); 
					while(noon.equals("")){
						noon = sheet.getCell(1, j--).getContents(); 
					}
					String time = sheet.getCell(2, i).getContents(); 
					String classHour = sheet.getCell(3, i).getContents(); 
					String trainingContent = sheet.getCell(4, i).getContents(); 
					String teacher = sheet.getCell(5, i).getContents(); 
					String place = sheet.getCell(6, i).getContents(); 
					String insertSql = "  insert into 培训课程表(日期,时段,时间,学时,培训内容,教师,地点,培训自编号,培训时间,期数) values(#"
					+dayTime+"#,'"+noon+"','"+time+"','"+classHour+"','"+trainingContent+"','"+teacher+"','"+place+"','"+theform.getTrainingID()+"','"+theform.getTrainingTime()+"','"+theform.getPeriod()+"')";
					stmt.execute(insertSql);
				}
			} catch (Exception e) {
				e.printStackTrace();
				conn.rollback();
				returnInfo.setFlag(false);
				return returnInfo;
			}  
		conn.commit();
        stmt.close();  
        conn.close(); 	
		returnInfo.setFlag(true);
		return returnInfo;
	}
	//教师信息表
	public ReturnInfo importteacherinfoExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();			
				for(int i=3;i<rowSize;i++){	
					String className = sheet.getCell(1, i).getContents(); 
					if(className.trim().equals("")){
						break;
					}
					String trainingID = sheet.getCell(2, i).getContents(); 
					String rainingContent = sheet.getCell(3, i).getContents(); 
					String teacherName = sheet.getCell(4, i).getContents(); 
					String sex = sheet.getCell(5, i).getContents(); 
					String teacherNum = sheet.getCell(6, i).getContents(); 
					String teacherType = sheet.getCell(7, i).getContents(); 
					String teacherTitle = sheet.getCell(8, i).getContents(); 
					String workUnit = sheet.getCell(9, i).getContents(); 
					String classHour = sheet.getCell(10, i).getContents(); 
					String remunerationStandard = sheet.getCell(11, i).getContents(); 
					String afterTaxExpenses = sheet.getCell(12, i).getContents(); 
					String OfficePhone = sheet.getCell(13, i).getContents(); 
					String mobile = sheet.getCell(14, i).getContents(); 
					String unit = sheet.getCell(15, i).getContents(); 
					String education = sheet.getCell(16, i).getContents(); 
					String professional = sheet.getCell(17, i).getContents(); 
					String email = sheet.getCell(18, i).getContents(); 
					String remark = sheet.getCell(19, i).getContents(); 
					String insertSql = "  insert into 教师信息表(培训班名称,培训项目编号,培训内容,教师姓名,性别,教师编码,教师分类,教师职称,工作单位名称及职务,课时,酬金标准,税后费用,办公电话,移动电话,直接聘请单位或机构名称,最新学历,专业方向,email,备注,培训时间,期数) values('"
					+className+"','"+trainingID+"','"+rainingContent+"','"+teacherName+"','"+sex+"','"+teacherNum+"','"+teacherType+"','"+teacherTitle+"','"+workUnit+"','"
					+classHour+"','"+remunerationStandard+"','"+afterTaxExpenses+"','"+OfficePhone+"','"+mobile+"','"+unit+"','"+education+"','"+professional+"','"+email+"','"+remark+"','"+theform.getTrainingTime()+"','"+theform.getPeriod()+"')";
					stmt.execute(insertSql);
				}
			} catch (Exception e) {
				conn.rollback();
				e.printStackTrace();			
				returnInfo.setFlag(false);
				return returnInfo;
			}
		conn.commit();
        stmt.close();  
        conn.close(); 				 
		returnInfo.setFlag(true);
		return returnInfo;
	}	
	public ReturnInfo importbookinfoExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();			
				for(int i=4;i<rowSize;i++){	
					String teachingMaterial = sheet.getCell(1, i).getContents(); 
					if(teachingMaterial.trim().equals("")){
						break;
					}
					String editor = sheet.getCell(2, i).getContents(); 
					String publishingUnit = sheet.getCell(3, i).getContents(); 
					String publishingTime = sheet.getCell(4, i).getContents(); 
					String bookNumber = sheet.getCell(5, i).getContents(); 
					String price = sheet.getCell(6, i).getContents(); 
					String number = sheet.getCell(7, i).getContents(); 
					String time = sheet.getCell(8, i).getContents(); 
					String remark = sheet.getCell(9, i).getContents(); 
					String insertSql = "  insert into 教材信息表(培训项目编号,培训班名称,教材名称,主编,出版单位,出版时间,书刊号,定价,订购册数,备注,培训时间,期数) values('"+theform.getTrainingID()+"','"
					+teachingMaterial+"','"+editor+"','"+publishingUnit+"','"+publishingTime+"','"+bookNumber+"','"+price+"','"+number+"','"+time+"','"+remark+"','"+theform.getTrainingTime()+"','"+theform.getPeriod()+"')";
					stmt.execute(insertSql);
				}
			} catch (Exception e) {
				conn.rollback();
				e.printStackTrace();			
				returnInfo.setFlag(false);
				return returnInfo;
			}
		conn.commit();
        stmt.close();  
        conn.close(); 				 
		returnInfo.setFlag(true);
		return returnInfo;
	}	
	public ReturnInfo importStudentinfoExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();			
				for(int i=2;i<rowSize;i++){	
/*					String teachingMaterial = sheet.getCell(1, i).getContents(); 
					if(teachingMaterial.trim().equals("")){
						break;
					}
					String editor = sheet.getCell(2, i).getContents(); 
					String publishingUnit = sheet.getCell(3, i).getContents(); 
					String publishingTime = sheet.getCell(4, i).getContents(); 
					String bookNumber = sheet.getCell(5, i).getContents(); 
					String price = sheet.getCell(6, i).getContents(); 
					String number = sheet.getCell(7, i).getContents(); 
					String time = sheet.getCell(8, i).getContents(); 
					String remark = sheet.getCell(9, i).getContents(); */
					String insertSql = "  insert into 学员信息表(培训班编号,期数,单位,姓名,性别,开票单位,岗位,手机,员工编号,用工类别,缴费方式) values('"+theform.getTrainingID()+"','"+theform.getPeriod()+"','"
					+sheet.getCell(1, i).getContents()+"','"+sheet.getCell(2, i).getContents()+"','"+sheet.getCell(3, i).getContents()+"','"+sheet.getCell(4, i).getContents()+"','"+sheet.getCell(5, i).getContents()+"','"+
					sheet.getCell(6, i).getContents()+"','"+sheet.getCell(7, i).getContents()+"','"+sheet.getCell(8, i).getContents()+"','"+sheet.getCell(9, i).getContents()+"')";
					stmt.execute(insertSql);
				}
			} catch (Exception e) {
				conn.rollback();
				e.printStackTrace();			
				returnInfo.setFlag(false);
				return returnInfo;
			}
		conn.commit();
        stmt.close();  
        conn.close(); 				 
		returnInfo.setFlag(true);
		return returnInfo;
	}		
	//收费信息
	public ReturnInfo importChargeinfoExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();			
				for(int i=3;i<rowSize;i++){	
					String name = sheet.getCell(1, i).getContents(); 
					String chargeType = sheet.getCell(5, i).getContents();
					String trainingTime = sheet.getCell(6, i).getContents(); 
					//String trainingID = sheet.getCell(7, i).getContents(); 
					String fee = sheet.getCell(10, i).getContents(); 
					String projectWorkPeople = sheet.getCell(11, i).getContents(); 		
					String insertSql = " update 学员信息表  set 培训通知月份 ='"+trainingTime+"',培训费='"+fee+"',缴费方式='"+chargeType+"',运行专责='"+projectWorkPeople+"'"
					+" where 培训zd编号='"+theform.getTrainingID()+"' and 期数='"+theform.getPeriod()
					+"' and 姓名 = '"+name+"';";
					stmt.execute(insertSql);
				}
			} catch (Exception e) {
				conn.rollback();
				e.printStackTrace();			
				returnInfo.setFlag(false);
				return returnInfo;
			}
		conn.commit();
        stmt.close();  
        conn.close(); 				 
		returnInfo.setFlag(true);
		return returnInfo;
	}		
	public ReturnInfo importFeeinfoExcel(PowerForm theform, HttpSession session) throws FileNotFoundException, IOException, SQLException{
		ReturnInfo returnInfo = new ReturnInfo();
		Workbook wb = null;
		int erro=0;

	
			// *1 读取Excel表数据
			InputStream new_is = theform.getUploadFile().getInputStream();
			
			//String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
			Connection conn = DriverManager.getConnection(Constant.getDburl(), "username", "password");  
			Statement stmt = conn.createStatement(); 
		    try {
		    	wb = Workbook.getWorkbook(new_is);
				Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");			      
			    conn.setAutoCommit(false);
			    
				// 扫描第一个sheet
				Sheet sheet = wb.getSheet(0); 
				int rowSize=sheet.getRows();			
				for(int i=1;i<rowSize;i++){	
					String updateSql = " update 培训班基础表  set 外聘教师酬金 ='"+sheet.getCell(3, i).getContents()+"',外聘教师差旅费='"+sheet.getCell(4, i).getContents()
							+"',课堂服务费='"+sheet.getCell(5, i).getContents()+"',教师食宿费='"+sheet.getCell(6, i).getContents()+"',午餐费='"+sheet.getCell(7, i).getContents()
							+"',实习参观费='"+sheet.getCell(8, i).getContents()+"',教材及资料费='"+sheet.getCell(9, i).getContents()+"',培训耗材='"+sheet.getCell(10, i).getContents()
							+"',委外培训费='"+sheet.getCell(11, i).getContents()
							+"' where 培训项目编号='"+theform.getTrainingID()+"' and 期数='"+theform.getPeriod()+"'";
					stmt.execute(updateSql);
				}
			} catch (Exception e) {
				conn.rollback();
				e.printStackTrace();			
				returnInfo.setFlag(false);
				return returnInfo;
			}
		conn.commit();
        stmt.close();  
        conn.close(); 				 
		returnInfo.setFlag(true);
		return returnInfo;
	}			
}

