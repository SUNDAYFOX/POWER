package power;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import javax.servlet.http.HttpServletRequest;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ToExcle {
	private boolean saveOnExit;  
    /** 
     * word�ĵ� 
     */  
    Dispatch doc = null;  
     
    /** 
     * word���г������s 
     */  
    private   ActiveXComponent word;  
    /** 
     * ����word�ĵ� 
     */  
    private   Dispatch documents;  
     
     
    /** 
     * ���캯�� 
     */  
    public ToExcle() {  
        if(word==null){  
        word = new ActiveXComponent("Word.Application");  
        word.setProperty("Visible",new Variant(false));  
        }  
        if(documents==null)  
        documents = word.getProperty("Documents").toDispatch();  
        saveOnExit = false;  
    }  
     
    /** 
     * ���ò������˳�ʱ�Ƿ񱣴� 
     * @param saveOnExit boolean true-�˳�ʱ�����ļ���false-�˳�ʱ�������ļ� 
     */  
    public void setSaveOnExit(boolean saveOnExit) {  
        this.saveOnExit = saveOnExit;  
    }  
    /** 
     * �õ��������˳�ʱ�Ƿ񱣴� 
     * @return boolean true-�˳�ʱ�����ļ���false-�˳�ʱ�������ļ� 
     */  
    public boolean getSaveOnExit() {  
        return saveOnExit;  
    }  
     
    /** 
     * ���ļ� 
     * @param inputDoc String Ҫ�򿪵��ļ���ȫ·�� 
     * @return Dispatch �򿪵��ļ� 
     */  
    public Dispatch open(String inputDoc) {  
        return Dispatch.call(documents,"Open",inputDoc).toDispatch();  
    }  
     
    /** 
     * ѡ������ 
     * @return Dispatch ѡ���ķ�Χ������ 
     */  
    public Dispatch select() {  
        return word.getProperty("Selection").toDispatch();  
    }  
     
    /** 
     * ��ѡ�����ݻ����������ƶ� 
     * @param selection Dispatch Ҫ�ƶ������� 
     * @param count int �ƶ��ľ��� 
     */  
    public void moveUp(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveUp");  
    }  
     
    /** 
     * ��ѡ�����ݻ����������ƶ� 
     * @param selection Dispatch Ҫ�ƶ������� 
     * @param count int �ƶ��ľ��� 
     */  
    public void moveDown(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveDown");  
    }  
     
    /** 
     * ��ѡ�����ݻ����������ƶ� 
     * @param selection Dispatch Ҫ�ƶ������� 
     * @param count int �ƶ��ľ��� 
     */  
    public void moveLeft(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++) {  
            Dispatch.call(selection,"MoveLeft");  
        }  
    }  
     
    /** 
     * ��ѡ�����ݻ����������ƶ� 
     * @param selection Dispatch Ҫ�ƶ������� 
     * @param count int �ƶ��ľ��� 
     */  
    public void moveRight(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveRight");  
    }  
     
    /** 
     * �Ѳ�����ƶ����ļ���λ�� 
     * @param selection Dispatch ����� 
     */  
    public void moveStart(Dispatch selection) {  
        Dispatch.call(selection,"HomeKey",new Variant(6));  
    }  
     
    /** 
     * ��ѡ�����ݻ����㿪ʼ�����ı� 
     * @param selection Dispatch ѡ������ 
     * @param toFindText String Ҫ���ҵ��ı� 
     * @return boolean true-���ҵ���ѡ�и��ı���false-δ���ҵ��ı� 
     */  
    public boolean find(Dispatch selection,String toFindText) {  
        //��selection����λ�ÿ�ʼ��ѯ  
        Dispatch find = word.call(selection,"Find").toDispatch();  
        //����Ҫ���ҵ�����  
        Dispatch.put(find,"Text",toFindText);  
        //��ǰ����  
        Dispatch.put(find,"Forward","True");  
        //���ø�ʽ  
        Dispatch.put(find,"Format","True");  
        //��Сдƥ��  
        Dispatch.put(find,"MatchCase","True");  
        //ȫ��ƥ��  
        Dispatch.put(find,"MatchWholeWord","True");  
        //���Ҳ�ѡ��  
        return Dispatch.call(find,"Execute").getBoolean();  
    }  
     
    /** 
     * ��ѡ�������滻Ϊ�趨�ı� 
     * @param selection Dispatch ѡ������ 
     * @param newText String �滻Ϊ�ı� 
     */  
    public void replace(Dispatch selection,String newText) {  
        //�����滻�ı�  
        Dispatch.put(selection,"Text",newText);  
    }  
     
    /** 
     * ȫ���滻 
     * @param selection Dispatch ѡ�����ݻ���ʼ����� 
     * @param oldText String Ҫ�滻���ı� 
     * @param newText String �滻Ϊ�ı� 
     */  
    public void replaceAll(Dispatch selection,String oldText,Object replaceObj) {  
        //�ƶ����ļ���ͷ  
        moveStart(selection);  
         
        if(oldText.startsWith("table") || replaceObj instanceof ArrayList)  
            replaceTable(selection,oldText,(ArrayList) replaceObj);  
        else {  
            String newText = (String) replaceObj;  
            if(newText==null)  
                newText="";  
            if(oldText.indexOf("image") != -1&!newText.trim().equals("") || newText.lastIndexOf(".bmp") != -1 || newText.lastIndexOf(".jpg") != -1 || newText.lastIndexOf(".gif") != -1){  
                while(find(selection,oldText)) {  
                    replaceImage(selection,newText);  
                    Dispatch.call(selection,"MoveRight");  
                }  
            }else{  
                while(find(selection,oldText)) {  
                    replace(selection,newText);  
                    Dispatch.call(selection,"MoveRight");  
                }  
            }  
        }  
    }  
     
    /** 
     * �滻ͼƬ 
     * @param selection Dispatch ͼƬ�Ĳ���� 
     * @param imagePath String ͼƬ�ļ���ȫ·���� 
     */  
    public void replaceImage(Dispatch selection,String imagePath) {  
        Dispatch.call(Dispatch.get(selection,"InLineShapes").toDispatch(),"AddPicture",imagePath);  
    }  
     
    /** 
     * �滻��� 
     * @param selection Dispatch ����� 
     * @param tableName String ������ƣ� 
     * ����table$1@1��table$2@1...table$R@N��R����ӱ���еĵ�N�п�ʼ��䣬N����word�ļ��еĵ�N�ű� 
     * @param fields HashMap �����Ҫ�滻���ֶ������ݵĶ�Ӧ�� 
     */  
    public void replaceTable(Dispatch selection,String tableName,ArrayList dataList) {  
        if(dataList.size() <= 1) {  
            System.out.println("Empty table!");  
            return;  
        }  
         
        //Ҫ������  
        String[] cols = (String[]) dataList.get(0);  
         
        //������  
        String tbIndex = tableName.substring(tableName.lastIndexOf("@") + 1);  
        //�ӵڼ��п�ʼ���  
        int fromRow = Integer.parseInt(tableName.substring(tableName.lastIndexOf("$") + 1,tableName.lastIndexOf("@")));  
        //���б��  
        Dispatch tables = Dispatch.get(doc,"Tables").toDispatch();  
        //Ҫ���ı��  
        Dispatch table = Dispatch.call(tables,"Item",new Variant(tbIndex)).toDispatch();  
        //����������  
        Dispatch rows = Dispatch.get(table,"Rows").toDispatch();  
        //�����  
        for(int i = 1;i < dataList.size();i ++) {  
            //ĳһ������  
            String[] datas = (String[]) dataList.get(i);  
             
            //�ڱ�������һ��  
            if(Dispatch.get(rows,"Count").getInt() < fromRow + i - 1)  
                Dispatch.call(rows,"Add");  
            //�����е������  
            for(int j = 0;j < datas.length;j ++) {  
                //�õ���Ԫ��  
                Dispatch cell = Dispatch.call(table,"Cell",Integer.toString(fromRow + i - 1),cols[j]).toDispatch();  
                //ѡ�е�Ԫ��  
                Dispatch.call(cell,"Select");  
                //���ø�ʽ  
                Dispatch font = Dispatch.get(selection,"Font").toDispatch();  
                Dispatch.put(font,"Bold","0");  
                Dispatch.put(font,"Italic","0");  
                //��������  
                Dispatch.put(selection,"Text",datas[j]);  
            }  
        }  
    }  
     
    /** 
     * �����ļ� 
     * @param outputPath String ����ļ�������·���� 
     */  
    public void save(String outputPath) {  
        Dispatch.call(Dispatch.call(word,"WordBasic").getDispatch(),"FileSaveAs",outputPath);  
    }  
     
    /** 
     * �ر��ļ� 
     * @param document Dispatch Ҫ�رյ��ļ� 
     */  
    public void close(Dispatch doc) {  
        Dispatch.call(doc,"Close",new Variant(saveOnExit));  
        word.invoke("Quit",new Variant[]{});  
        word = null;  
    }  
     
    /** 
     * ����ģ�塢��������word�ļ� 
     * @param inputPath String ģ���ļ�������·���� 
     * @param outPath String ����ļ�������·���� 
     * @param data HashMap ���ݰ�������Ҫ�����ֶΡ���Ӧ�����ݣ� 
     */  
    public void toWord(String inputPath,String outPath,HashMap data) {  
        String oldText;  
        Object newValue;  
        try {  
            if(doc==null)  
            doc = open(inputPath);  
             
            Dispatch selection = select();  
             
            Iterator keys = data.keySet().iterator();  
            while(keys.hasNext()) {  
                oldText = (String) keys.next();  
                newValue = data.get(oldText);  
                 
                replaceAll(selection,oldText,newValue);  
            }  
             
            save(outPath);  
        } catch(Exception e) {  
            System.out.println("����word�ļ�ʧ�ܣ�");  
            e.printStackTrace();  
        } finally {  
            if(doc != null)  
                close(doc);  
        }  
    }  
     
    public synchronized static void word(String inputPath,String outPath,HashMap data){  
    	ToExcle j2w = new ToExcle();  
        j2w.toWord(inputPath,outPath,data);  
    }  
     
    @SuppressWarnings({ "rawtypes", "unchecked" })  
   public Boolean schoolclass(PowerForm theform,HttpServletRequest request) throws SQLException {  
    	Boolean res = true;
    	HashMap data = new HashMap();
	    try {
			Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
	    Connection conn;
		conn = DriverManager.getConnection(dbur1, "username", "password");
		Statement stmt = conn.createStatement();  
		Statement stmt2 = conn.createStatement();  
		String sql1 = ("select ��ѵ�Ա��,��Ŀ����,��ѵ����,��Ŀ��������,��Ŀ����������,���������˵绰,��Ŀ�ֿ�,��Ŀ����������,���������˵绰  from ������ѵ�ƻ���  where ��ѵ�Ա�� = '"+theform.getTrainingID()+"'");  
	    String sql2 = ("select a.����,a.����ʱ��,a.����ʱ��,a.��ѵ�ص�,a.�ص�,a.������ ,b.�칫�绰, b.�ƶ��绰  from ��ѵ�����б�  a left join ��ʦ��Ϣ�� b on a.������=b.��ʦ����"
	    		+ " where a.��ѵ��Ŀ��� = '"+theform.getTrainingID()+"'");  
	    ResultSet rs1 = stmt.executeQuery(sql1);  
	    ResultSet rs2 = stmt2.executeQuery(sql2);  
	    //stmt.execute(insertSql);
	    SimpleDateFormat format = new SimpleDateFormat("yyyy��MM��dd");
	    SimpleDateFormat format2 = new SimpleDateFormat("dd��");
    	data.put("$trainingid$","");
    	data.put("$classname$","");
    	data.put("$object$","");
    	data.put("$projectDevelopmentDept$","");
    	data.put("$projectDevelopmentPeople$","");
    	data.put("$projectPeoplePhone$","");
		data.put("$projectType$","");
		data.put("$verify$","");
		data.put("$projectworkPeople$","");
		data.put("$projectworkPeoplePhone$","");
    	data.put("$num$","");
    	data.put("$trainingtime$","");
    	data.put("$time$","");
    	data.put("$place$","");
    	data.put("$trainingplace$","");
    	data.put("$teacher1$","");
    	data.put("$teacher2$","");	   
    	data.put("$teacherphone1$","");
    	data.put("$teacherphone2$","");	   
	    while (rs1.next()) {  
	    	data.put("$trainingid$",rs1.getString(1));
	    	data.put("$classname$",rs1.getString(2));
	    	data.put("$object$",rs1.getString(3));
	    	data.put("$projectDevelopmentDept$",rs1.getString(4));
	    	data.put("$projectDevelopmentPeople$",rs1.getString(5));
	    	data.put("$projectPeoplePhone$",rs1.getString(6));
	    	String projectType = rs1.getString(7);
	    	if((projectType!=null&&projectType.equals("����"))){
	    		data.put("$projectType$","����������ѵ��");
	    		data.put("$verify$","����");
	    	}else if(projectType!=null&&projectType.equals("����")){
	    		data.put("$projectType$","������ѵ��");
	    		data.put("$verify$","������");
	    	}else{
	    		data.put("$projectType$","");
	    		data.put("$verify$","");	    		
	    	}
	    	data.put("$projectworkPeople$",rs1.getString(8));
	    	String projectworkPeoplePhone = rs1.getString(9);
	    	data.put("$projectworkPeoplePhone$",projectworkPeoplePhone.replace("/", "��"));
	    }  
	    while (rs2.next()) {  
	    	data.put("$num$",rs2.getString(1));
	    	Date startTime = rs2.getDate(2);
	    	Date endTime = rs2.getDate(3);
	    	String time = format.format(startTime)+"-"+format2.format(endTime);
	    	data.put("$trainingtime$",time);
	    	data.put("$time$",startTime+"��8��00ǰ");
	    	String place = rs2.getString(4);
	    	if(place.equals("����ƺ"))
	    	{
	    		place = "�������켼�����ı���������ƺ���д�����̨";
	    	}else if(place.equals("ɳƺ��")){
	    		place = "�������켼�����ķֲ���ɳƺ�ӣ��д�����̨";
	    	}else if(place.equals("����ɽ")){
	    		place = "�������켼�����ķֲ�������ɽ���д�����̨";
	    	}
	    	data.put("$place$",place);
	    	data.put("$trainingplace$",rs2.getString(5));
	    	String teacher = rs2.getString(6);
	    	String teacherPhone = rs2.getString(7);
	    	String teacherMobile = rs2.getString(8);
	    	if(teacher.indexOf("\\")>-1){
		    	data.put("$teacher1$",teacher.split("\\")[0]);
		    	data.put("$teacher2$",teacher.split("\\")[1]);	    		
	    	}else{
		    	data.put("$teacher1$",teacher);
		    	data.put("$teacher2$","");	
		    	if(teacherPhone == null){
		    		teacherPhone = "";
		    	}
		    	if(teacherMobile == null){
		    		teacherMobile = "";
		    	}
		    	data.put("$teacherphone1$",teacherPhone+"��"+teacherMobile);
		    	data.put("$teacherphone2$","");	   	
		    	
	    	}

	    	//data.put("$remark$",rs2.getString(6));
	    	//data.put("$classname$",rs1.getString(2));
	    	//data.put("$object$",rs1.getString(3));
	    }  
	    rs1.close();  
	    stmt.close();  
	    conn.close();            
        ToExcle j2w = new ToExcle();  
        long time1 = System.currentTimeMillis(); 
        System.out.println(request.getRealPath("")+"\\DB\\template\\У�ڿ���֪ͨ.doc"); 
        ComThread.InitSTA();
        j2w.toWord(request.getRealPath("")+"\\DB\\template\\У�ڿ���֪ͨ.doc",request.getRealPath("")+"\\DB\\result\\schoolclass.doc",data);  
        ComThread.Release(); 
        System.out.println("time cost : " + (System.currentTimeMillis() - time1));  
        
        return res;
    }  	

}
