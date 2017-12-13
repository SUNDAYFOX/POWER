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
    	FileItem item=(FileItem)it.next();//ÿһ��item�ʹ���һ���������

    	if(item.isFormField()){//�ж�input�Ƿ�Ϊ��ͨ��������

    	String name=item.getFieldName();

    	String value= item.getString();
    	}else{

    	//�õ��ϴ��ļ�������,����ȡ
    	filename=item.getName();

    	//�õ��ϴ��ļ�Ҫд���Ŀ¼
    	path=this.getServletContext().getRealPath("\\DB");
    	//����Ŀ¼���ļ����������
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
    //��˾������ѵ�ƻ�����
     public static void PlanInDB(String path,String filename) throws ClassNotFoundException, SQLException{
 		Map<String,List> classmap = new HashMap<String,List>();
 		Map<String,List> projectmap = new HashMap<String,List>();
	    Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");  
	    /** 
	     * ֱ������access�ļ��� 
	     */  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
	    Connection conn = DriverManager.getConnection(dbur1, "username", "password");  
	    Statement stmt = conn.createStatement();  
 		try {

 			// ��������������ȡExcel  
 			InputStream is = new FileInputStream(new File(path+"\\"+filename).getAbsolutePath());  
 			// jxl�ṩ��Workbook��  
 			Workbook wbin = Workbook.getWorkbook(is);  
 			// Excel��ҳǩ����  
 			int sheet_size = wbin.getNumberOfSheets();  
 			//for (int index = 0; index < sheet_size; index++) {  
 			// ÿ��ҳǩ����һ��Sheet����  
 			Sheet sheet = wbin.getSheet(0);  
 			// sheet.getRows()���ظ�ҳ��������  
 			for (int i = 2; i < sheet.getRows(); i++) {  
 				String trainingID = sheet.getCell(1, i).getContents(); 
 				String companyID = sheet.getCell(2, i).getContents();
 				String projectID = sheet.getCell(3, i).getContents();
 				String projectName = sheet.getCell(4, i).getContents();
 				String trainingNum = sheet.getCell(6, i).getContents();
 				String name = sheet.getCell(21, i).getContents(); 
 				String insertSql = "  insert into ������ѵ�ƻ���(��ѵ�Ա��,��ҵ���,��Ŀ����,��Ŀ���,��ѵʱ��,��ѵ����,���� ,��Ŀ����������) values('"+trainingID+"','"+companyID+"','"+projectName+"','"+projectID+"','";
				String months = sheet.getCell(11, i).getContents().replace("��", "");
				String month[]  =months.split("��"); 
				if(trainingID.equals("jj2017-01-94")){
					System.out.println(projectName);
				}
				if(months.indexOf("��")>-1){//5��11��
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
				}else{//5(2)��
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
 					int total = Integer.parseInt(sheet.getCell(6, i).getContents())+ list.get(0);//������
 					list.set(0, total);
 					projectlist.set(0, projectlist.get(0)+1);
 					String months = sheet.getCell(11, i).getContents().replace("��", "");
 					String month[]  =months.split("��"); 
 					if(months.indexOf("��")>-1){//5��11��
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
 					}else{//5(2)��
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
 					list.set(0, total);//�༶��
 					projectlist.set(0, 1);//��Ŀ��
 					String months = sheet.getCell(11, i).getContents().replace("��", "");
 					String month[]  =months.split("��"); 

 					if(months.indexOf("��")>-1){//5��11��
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
 					}else{//5(2)��
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
 			 classmap.put("�ϼ�", classtotallist);
 			 projectmap.put("�ϼ�", projecttotallist);	
 			wbin.close();
 			File f = new File(path+"\\classanalysis.xls");
 			Workbook wbout = Workbook.getWorkbook(f);//
 			WritableWorkbook book = wbout.createWorkbook(f, wbout);
 			// Sheet sheet = wb.getSheet(0); // ��õ�һ�����������
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
 						jxl.format.CellFormat cf = cel.getCellFormat();//��ȡ��һ����Ԫ��ĸ�ʽ  
 						int content=0;
 						if(i<14){
 							content = classcontent;
 						}else{
 							content = projectcontent;
 						}
 						Label lbl = new jxl.write.Label(j, i, String.valueOf(content));//����һ����Ԫ���ֵ��Ϊ���޸����ֵ��  
 			            lbl.setCellFormat(cf);//���޸ĺ�ĵ�Ԫ��ĸ�ʽ�趨�ɸ�ԭ��һ�� 
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
