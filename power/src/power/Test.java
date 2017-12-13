package power;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class Test {
	public static void main(String[] args) {
/*		try { // 防止文件建立或读取失败，用catch捕捉错误并打印，也可以throw  

			         读入TXT文件   
        String pathname = "F:\\TOOL\\tomcat-6.0.14\\DB\\123.txt"; 
        File filename = new File(pathname);
        InputStreamReader reader = new InputStreamReader(  
        new FileInputStream(filename)); // 建立一个输入流对象reader  
        BufferedReader br = new BufferedReader(reader); // 建立一个对象，它把文件内容转成计算机能读懂的语言  
        String line = "";  
        line = br.readLine();  
        Map<String,List> map = new HashMap<String,List>();
       while (line != null) {  
            line = br.readLine(); // 一次读入一行数据  
            String str[] =line.split("|"); 
            List list = new ArrayList();
            System.out.println(line);
            if(line.split("|")[3].equals("安全知识综合业务调考")){
            	System.out.println(line);
            }
            String months = str[4];


            if(months.indexOf("、")>-1){//5、11月
            	String month[]  =months.split("、"); 
            	for (int i = 0; i < month.length; i++) {
					if(month[i].indexOf("(")>-1){
						int num = Integer.parseInt((String) list.get(Integer.parseInt(month[i].substring(0, month[i].indexOf("(")))));
						num= num+Integer.parseInt(month[i].substring(month[i].indexOf("("), month[i].indexOf(")")));
						list.add(Integer.parseInt(month[i].substring(0, month[i].indexOf("("))), num);
					}else{
						int num = Integer.parseInt((String) list.get(Integer.parseInt(month[i])));
						num= num+Integer.parseInt(month[i]);
						list.add(Integer.parseInt(month[i]), num);						
					}
				}
            	//list.add(Integer.parseInt(s), element);
            }else{//5(2)月

            }

            map.put("str[14]", list);
        }  

			         写入Txt文件   
        File writename = new File(".\\result\\en\\output.txt"); // 相对路径，如果没有则要建立一个新的output。txt文件  
        writename.createNewFile(); // 创建新文件  
        BufferedWriter out = new BufferedWriter(new FileWriter(writename));  
        out.write("我会写入文件啦\r\n"); // \r\n即为换行  
        out.flush(); // 把缓存区内容压入文件  
        out.close(); // 最后记得关闭文件  
			 
		} catch (Exception e) {  
			e.printStackTrace();  
		}  */
		Map<String,List> classmap = new HashMap<String,List>();
		Map<String,List> projectmap = new HashMap<String,List>();
		try {

			// 创建输入流，读取Excel  
			InputStream is = new FileInputStream(new File("D:\\TOOL\\tomcat-6.0.14\\DB\\123.xls").getAbsolutePath());  
			// jxl提供的Workbook类  
			Workbook wbin = Workbook.getWorkbook(is);  
			// Excel的页签数量  
			int sheet_size = wbin.getNumberOfSheets();  
			//for (int index = 0; index < sheet_size; index++) {  
			// 每个页签创建一个Sheet对象  
			Sheet sheet = wbin.getSheet(1);  
			// sheet.getRows()返回该页的总行数  
			for (int i = 2; i < sheet.getRows(); i++) {  
				String name = sheet.getCell(21, i).getContents(); 
				if(sheet.getCell(1, i).getContents().equals("")){
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
			}  
			List classtotallist = new ArrayList<Integer>();
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
			File f = new File("d:\\TOOL\\tomcat-6.0.14\\DB\\班期数分析表.xls");
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

			}
			book.write();
			book.close();
			wbout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
