package upload;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileUploadException;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;  
public class Upload extends HttpServlet {  
	  
    private static final long serialVersionUID = 6777945010008132796L;  
  
  @Override  
  protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {  
      super.doGet(req, resp);  
      System.out.println("doGet");  
      resp.sendRedirect("/upload/index.jsp");  
  }  
    
    @Override  
    protected void doPost(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {  
    	DiskFileItemFactory factory=new DiskFileItemFactory();

    	ServletFileUpload upload=new ServletFileUpload(factory);
    	 String username = req.getParameter( "username" );
    	String path="";
    	String filename="";
    	try
    	{
    	List list=upload.parseRequest(req);
    	Iterator it=list.iterator();

    	while(it.hasNext())
    	{
    	FileItem item=(FileItem)it.next();//每一个item就代表一个表单输出项

    	if(item.isFormField()){//判断input是否为普通表单输入项

    	String name=item.getFieldName();

    	String value= item.getString();
    	}else{

    	//得到上传文件的名称,并截取
    	filename=item.getName();

    	//得到上传文件要写入的目录
    	path=this.getServletContext().getRealPath("\\DB");
    	//根据目录和文件创建输出流
       // String xxx = "d:\\"+filename;
    	FileOutputStream out=new FileOutputStream(path+"\\"+filename);

    	InputStream in = item.getInputStream();
    	 byte buffer[] = new byte[1024];
    	 int len = 0;
    	while((len=in.read(buffer))>0){
    	 out.write(buffer,0,len);
    	   }
    	in.close();
    	out.close();
    	 }
    	}
    	} catch (FileUploadException e) 
    	{
    	e.printStackTrace();
    	  }
    	try {
			Upload.PlanInDB(path,filename);
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
        resp.sendRedirect("/power/index.jsp");     

    }  
    //公司短期培训计划导入
     public static void PlanInDB(String path,String filename) throws ClassNotFoundException, SQLException{
 		Map<String,List> classmap = new HashMap<String,List>();
 		Map<String,List> projectmap = new HashMap<String,List>();
	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	    /** 
	     * 直接连接access文件。 
	     */  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
 		try {

 			// 创建输入流，读取Excel  
 			InputStream is = new FileInputStream(new File(path+"\\"+filename).getAbsolutePath());  
 			// jxl提供的Workbook类  
 			Workbook wbin = Workbook.getWorkbook(is);  
 			// Excel的页签数量  
 			int sheet_size = wbin.getNumberOfSheets();  
 			//for (int index = 0; index < sheet_size; index++) {  
 			// 每个页签创建一个Sheet对象  
 			Sheet sheet = wbin.getSheet(0);  
 			// sheet.getRows()返回该页的总行数  
 			for (int i = 2; i < sheet.getRows(); i++) {  
 				String trainingID = sheet.getCell(1, i).getContents(); 
 				String companyID = sheet.getCell(2, i).getContents();
 				String projectID = sheet.getCell(3, i).getContents();
 				String projectName = sheet.getCell(4, i).getContents();
 				String trainingNum = sheet.getCell(6, i).getContents();
 				String name = sheet.getCell(21, i).getContents(); 
 				String insertSql = "  insert into 短期培训计划表(培训自编号,企业编号,项目名称,项目编号,培训时间,培训期数,期数 ,项目开发责任人) values('"+trainingID+"','"+companyID+"','"+projectName+"','"+projectID+"','";
				String months = sheet.getCell(11, i).getContents().replace("月", "");
				String month[]  =months.split("、"); 
				if(trainingID.equals("jj2017-01-94")){
					System.out.println(projectName);
				}
				if(months.indexOf("、")>-1){//5、11月
					for (int i1 = 0; i1 < month.length; i1++) {
						if(month[i1].indexOf("(")>-1){
							int num = Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
							for (int i2 = 1; i2 < num+1; i2++) {
								

								String sql = insertSql+month[i1].substring(0,month[i1].indexOf("("))+"','"+trainingNum+"','"+i2+"','"+name+"')";
								 stmt.execute(sql);
							}
						}else{
							String sql = insertSql+month[i1]+"','"+trainingNum+"','"+1+"','"+name+"')";
							 stmt.execute(sql);
						}
					}
				}else{//5(2)月
					if(months.indexOf("(")>-1){
						int num = Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
						for (int i3 = 1; i3 < num+1; i3++) {
							String sql = insertSql+months.substring(0,months.indexOf("("))+"','"+trainingNum+"','"+i3+"','"+name+"')";
							stmt.execute(sql);
						}
					}else{
						String sql = insertSql+months+"','"+trainingNum+"',"+1+",'"+name+"')";
						stmt.execute(sql);
					}     
				}
 			}
 			
/* 				if(sheet.getCell(1, i).getContents().equals("")){
 					System.out.println(name);
 					break;
 				}
 				System.out.println(sheet.getCell(1, i).getContents()); 
 				if(classmap.containsKey(name)){
 					List<Integer> list = classmap.get(name);
 					List<Integer> projectlist = projectmap.get(name);
 					int total = Integer.parseInt(sheet.getCell(6, i).getContents())+ list.get(0);//总期数
 					list.set(0, total);
 					projectlist.set(0, projectlist.get(0)+1);
 					String months = sheet.getCell(11, i).getContents().replace("月", "");
 					String month[]  =months.split("、"); 
 					if(months.indexOf("、")>-1){//5、11月
 						for (int i1 = 0; i1 < month.length; i1++) {
 							if(month[i1].indexOf("(")>-1){
 								int num =  list.get(Integer.parseInt(month[i1].substring(0, month[i1].indexOf("(")))-1);
 								projectlist.set(Integer.parseInt(month[i1].substring(0, month[i1].indexOf("(")))-1, num+1);
 								num= num+Integer.parseInt(month[i1].substring(month[i1].indexOf("(")+1, month[i1].indexOf(")")));
 								list.set(Integer.parseInt(month[i1].substring(0, month[i1].indexOf("(")))-1, num);
 							}else{
 								int num = list.get(Integer.parseInt(month[i1])-1);
 								list.set(Integer.parseInt(month[i1])-1, num+1);		
 								projectlist.set(Integer.parseInt(month[i1])-1, num+1);		
 							}
 						}
 						//list.add(Integer.parseInt(s), element);
 					}else{//5(2)月
 						if(months.indexOf("(")>-1){
 							int num= list.get(Integer.parseInt(months.substring(0,months.indexOf("(")))-1) ;//  0,months.substring(months.indexOf("(")));
 							projectlist.set(Integer.parseInt(months.substring(0, months.indexOf("(")))-1, num+1);
 							num= num+Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
 							list.set(Integer.parseInt(months.substring(0, months.indexOf("(")))-1, num);
 						}else{
 							int num= list.get(Integer.parseInt(months)-1);
 							list.set(Integer.parseInt(months)-1, num+1);
 							projectlist.set(Integer.parseInt(months)-1, num+1);
 						}     
 					}
 				}else{
 					List list  = new ArrayList();
 					List projectlist  = new ArrayList();
 					for (int j = 0; j < 12; j++) {
 						list.add(0);
 						projectlist.add(0);
 					}
 					int total = Integer.parseInt(sheet.getCell(6, i).getContents());
 					list.set(0, total);//班级数
 					projectlist.set(0, 1);//项目数
 					String months = sheet.getCell(11, i).getContents().replace("月", "");
 					String month[]  =months.split("、"); 

 					if(months.indexOf("、")>-1){//5、11月
 						for (int i1 = 0; i1 < month.length; i1++) {
 							if(month[i1].indexOf("(")>-1){
 								int num= Integer.parseInt(month[i1].substring(month[i1].indexOf("("), month[i1].indexOf(")")));
 								list.set(Integer.parseInt(month[i1].substring(0, month[i1].indexOf("("))), num);
 								projectlist.set(Integer.parseInt(month[i1].substring(0, month[i1].indexOf("("))), 1);
 							}else{
 								list.set(Integer.parseInt(month[i1])-1, 1);		
 								projectlist.set(Integer.parseInt(month[i1])-1, 1);		
 							}
 						}
 						//list.add(Integer.parseInt(s), element);
 					}else{//5(2)月
 						if(months.indexOf("(")>-1){
 							int num= Integer.parseInt(months.substring(months.indexOf("(")+1,months.indexOf(")")));
 							list.set(Integer.parseInt(months.substring(0, months.indexOf("(")))-1, num);
 							projectlist.set(Integer.parseInt(months.substring(0, months.indexOf("(")))-1, 1);
 						}else{
 							list.set(Integer.parseInt(months)-1, 1);		
 							projectlist.set(Integer.parseInt(months)-1, 1);		
 						}                     	
 					}

 					classmap.put(name, list);
 					projectmap.put(name, projectlist);
 				}
 			}  */
/* 			List classtotallist = new ArrayList<Integer>();
 			List projecttotallist = new ArrayList<Integer>();
 			for (int j = 0; j < 12; j++) {
 				 int classcount = 0;
 				 int projectcount = 0;
 				 for (String key : classmap.keySet()) {
 					   List<Integer> classcountlist = classmap.get(key);
 					   classcount = classcount+classcountlist.get(j);				   
 					  // System.out.println("key= "+ key + " and value= " + classmap.get(key));
 				  }		
 				 classtotallist.add(j,classcount);
 				 for (String key : projectmap.keySet()) {
 					   List<Integer> projectcountlist = projectmap.get(key);
 					   projectcount = projectcount+projectcountlist.get(j);
 					  
 					  // System.out.println("key= "+ key + " and value= " + projectmap.get(key));
 				  }		
 				  projecttotallist.add(j,projectcount);		 
 			}
 			 classmap.put("合计", classtotallist);
 			 projectmap.put("合计", projecttotallist);	
 			wbin.close();
 			File f = new File(path+"\\classanalysis.xls");
 			Workbook wbout = Workbook.getWorkbook(f);//
 			WritableWorkbook book = wbout.createWorkbook(f, wbout);
 			// Sheet sheet = wb.getSheet(0); // 获得第一个工作表对象
 			WritableSheet st = book.getSheet(0);
 			for (int i = 2; i < 26; i++) {
 				Cell name = st.getCell(0,i);
 				if(classmap.containsKey(name.getContents())){
 					for (int j = 1; j < st.getColumns(); j++) {
 						Cell cel = st.getCell(j,i);
 						List list = classmap.get(name.getContents());
 						List projectlist = projectmap.get(name.getContents());
 						int classcontent  = (Integer) list.get(j-1);
 						int projectcontent  = (Integer) projectlist.get(j-1);
 						//if(cel.getType() == CellType.LABEL){
 						jxl.format.CellFormat cf = cel.getCellFormat();//获取第一个单元格的格式  
 						int content=0;
 						if(i<14){
 							content = classcontent;
 						}else{
 							content = projectcontent;
 						}
 						Label lbl = new jxl.write.Label(j, i, String.valueOf(content));//将第一个单元格的值改为“修改後的值”  
 			            lbl.setCellFormat(cf);//将修改后的单元格的格式设定成跟原来一样 
 			            if(i==28){
 			            	System.out.println("i="+i+"j="+j);
 			            }
 						//jxl.write.Number numb = new jxl.write.Number(j, i, content); 
 						//Label label = new Label(j, i,String.valueOf(content) );
 			            if(cf!=null){
 			            	st.addCell(lbl);
 			            }
 						//st.insertColumn(1);
 						//}
 					}					
 				}

 			}*/
 			//book.write();
 			//book.close();
 			//wbout.close();
 		} catch (Exception e) {
 			e.printStackTrace();
			e.printStackTrace();
			conn.rollback();
 		}
     }
}  
