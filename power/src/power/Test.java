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
/*		try { // ��ֹ�ļ��������ȡʧ�ܣ���catch��׽���󲢴�ӡ��Ҳ����throw  

			         ����TXT�ļ�   
        String pathname = "F:\\TOOL\\tomcat-6.0.14\\DB\\123.txt"; 
        File filename = new File(pathname);
        InputStreamReader reader = new InputStreamReader(  
        new FileInputStream(filename)); // ����һ������������reader  
        BufferedReader br = new BufferedReader(reader); // ����һ�����������ļ�����ת�ɼ�����ܶ���������  
        String line = "";  
        line = br.readLine();  
        Map<String,List> map = new HashMap<String,List>();
       while (line != null) {  
            line = br.readLine(); // һ�ζ���һ������  
            String str[] =line.split("|"); 
            List list = new ArrayList();
            System.out.println(line);
            if(line.split("|")[3].equals("��ȫ֪ʶ�ۺ�ҵ�����")){
            	System.out.println(line);
            }
            String months = str[4];


            if(months.indexOf("��")>-1){//5��11��
            	String month[]  =months.split("��"); 
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
            }else{//5(2)��

            }

            map.put("str[14]", list);
        }  

			         д��Txt�ļ�   
        File writename = new File(".\\result\\en\\output.txt"); // ���·�������û����Ҫ����һ���µ�output��txt�ļ�  
        writename.createNewFile(); // �������ļ�  
        BufferedWriter out = new BufferedWriter(new FileWriter(writename));  
        out.write("�һ�д���ļ���\r\n"); // \r\n��Ϊ����  
        out.flush(); // �ѻ���������ѹ���ļ�  
        out.close(); // ���ǵùر��ļ�  
			 
		} catch (Exception e) {  
			e.printStackTrace();  
		}  */
		Map<String,List> classmap = new HashMap<String,List>();
		Map<String,List> projectmap = new HashMap<String,List>();
		try {

			// ��������������ȡExcel  
			InputStream is = new FileInputStream(new File("D:\\TOOL\\tomcat-6.0.14\\DB\\123.xls").getAbsolutePath());  
			// jxl�ṩ��Workbook��  
			Workbook wbin = Workbook.getWorkbook(is);  
			// Excel��ҳǩ����  
			int sheet_size = wbin.getNumberOfSheets();  
			//for (int index = 0; index < sheet_size; index++) {  
			// ÿ��ҳǩ����һ��Sheet����  
			Sheet sheet = wbin.getSheet(1);  
			// sheet.getRows()���ظ�ҳ��������  
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
			 classmap.put("�ϼ�", classtotallist);
			 projectmap.put("�ϼ�", projecttotallist);	
			wbin.close();
			File f = new File("d:\\TOOL\\tomcat-6.0.14\\DB\\������������.xls");
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

			}
			book.write();
			book.close();
			wbout.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
